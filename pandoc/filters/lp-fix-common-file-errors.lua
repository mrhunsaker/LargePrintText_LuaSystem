--[[
lp-fix-common-file-errors.lua

Port of the VistaType LP "Fix Common File Errors" macro
(VistaType LP User Guide, "File Cleanup Group", p.44-45).

IMPORTANT SCOPE NOTE
=====================
The original macro operates directly on Word's object model (Find/Replace
with wildcards across Range objects, paragraph style objects, inline shape
objects, etc.) It does 15+ distinct things. Pandoc's document model is a
structural AST, not a live word-processor object model, so not everything
translates. Honest breakdown:

  Handled in THIS filter (AST-level, deterministic, source-format-agnostic):
    - Collapse runs of repeated hyphen / en dash / em dash characters
      (each dash type collapses only with itself, per the guide's wording)
      down to a single instance, including inside mixed alphanumeric tokens
    - Strip a Space/Tab inline immediately before or after a text token
      that starts or ends with one of those dash characters
    - Collapse runs of literal space characters inside a single text run
    - Convert SmallCaps-styled text to plain all-caps text
    - Strip leading/trailing Tab and Space inlines from each paragraph
    - Convert a straight ' or " that immediately follows a digit into a
      proper prime (U+2032) / double prime (U+2033) mark

  NOT handled here (needs a different layer -- see docs/COMMAND_MAP.md):
    - Abbyy FineReader style remapping (needs original DOCX style IDs --
      see pandoc/preprocess/docx_style_remap.py, or do it live in
      LibreOffice: lo-macros/vistatype/lp/file_cleanup.py)
    - Drop caps -> normal text (OOXML "framePr" drop caps aren't modeled by
      Pandoc's AST at all; needs LibreOffice's real object model)
    - Un-boxing text boxes/frames while preserving their text (Pandoc's own
      DOCX reader already inlines *some* text boxes; anchored frames are not
      reliably handled -- do this in LibreOffice instead)
    - Removing in-line images with user confirmation (interactive,
      confirmation-gated -- implemented as an LO macro, not a batch filter)

Usage: chain with the other lp-*.lua filters. Run this one *before*
lp-tag-reference-page-numbers.lua so stray tabs/spaces don't interfere with
page-number pattern matching.
]]

local HYPHEN = "-"
local EN_DASH = "\226\128\147" -- U+2013
local EM_DASH = "\226\128\148" -- U+2014
local PRIME = "\226\128\178"        -- U+2032
local DOUBLE_PRIME = "\226\128\179" -- U+2033

local function dash_class(byte_seq)
    if byte_seq == HYPHEN then return "h" end
    if byte_seq == EN_DASH then return "n" end
    if byte_seq == EM_DASH then return "m" end
    return nil
end

-- Collapse runs of the *same* dash character (hyphen-run, en-dash-run, or
-- em-dash-run) down to a single instance. Operates on the whole token so it
-- also fixes dashes fused into a larger word, e.g. "triple---dash".
local function collapse_dash_runs(s)
    local out = {}
    local prev_class = nil
    for _, cp in utf8.codes(s) do
        local ch = utf8.char(cp)
        local cls = dash_class(ch)
        if not (cls and cls == prev_class) then
            table.insert(out, ch)
        end
        prev_class = cls
    end
    return table.concat(out)
end

local function convert_primes(s)
    s = s:gsub("(%d)'", "%1" .. PRIME)
    s = s:gsub('(%d)"', "%1" .. DOUBLE_PRIME)
    return s
end

local function starts_with_dash(s)
    local first = s:sub(1, 1)
    if first == HYPHEN then return true end
    local first3 = s:sub(1, 3)
    return first3 == EN_DASH or first3 == EM_DASH
end

local function ends_with_dash(s)
    if s:sub(-1) == HYPHEN then return true end
    local last3 = s:sub(-3)
    return last3 == EN_DASH or last3 == EM_DASH
end

function Str(el)
    el.text = collapse_dash_runs(el.text)
    el.text = el.text:gsub(" +", " ")
    el.text = convert_primes(el.text)
    return el
end

function SmallCaps(el)
    -- Flatten to plain text, upper-cased ("converts small caps to all caps")
    local plain = pandoc.utils.stringify(el)
    return pandoc.Str(string.upper(plain))
end

local function is_space_token(inline)
    -- Pandoc has no distinct "Tab" Inline type -- a literal tab character
    -- collapses to Space during parsing (verified: `printf 'a\tb'` parses
    -- to `Str "a", Space, Str "b"`). So Space is the only whitespace-run
    -- token we can see or remove here.
    return inline.t == "Space"
end

-- Remove a Space/Tab that sits directly next to a Str token which starts or
-- ends with a dash character, and strip leading/trailing whitespace from
-- the paragraph as a whole.
local function clean_inlines(inlines)
    local out = pandoc.List()
    local i = 1
    local n = #inlines
    while i <= n do
        local cur = inlines[i]
        if is_space_token(cur) then
            local prev = out[#out]
            local nxt = inlines[i + 1]
            local drop = false
            if prev and prev.t == "Str" and ends_with_dash(prev.text) then
                drop = true
            elseif nxt and nxt.t == "Str" and starts_with_dash(nxt.text) then
                drop = true
            end
            if not drop then
                out:insert(cur)
            end
        else
            out:insert(cur)
        end
        i = i + 1
    end
    -- strip leading/trailing whitespace tokens
    while #out > 0 and is_space_token(out[1]) do
        out:remove(1)
    end
    while #out > 0 and is_space_token(out[#out]) do
        out:remove(#out)
    end
    return out
end

function Para(el)
    el.content = clean_inlines(pandoc.List(el.content))
    return el
end

function Plain(el)
    el.content = clean_inlines(pandoc.List(el.content))
    return el
end

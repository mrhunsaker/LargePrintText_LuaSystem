--[[
lp-tag-reference-page-numbers.lua

Port of "Auto Tag Page Numbers" (VistaType LP User Guide p.47; BANA guide
p.16, same concept). A "reference page number" is a page number from the
*original print* document, preserved in the transcription so a reader/
teacher can cross-reference pages between the large-print (or braille) copy
and the original book -- not a physical page number of the output itself.

Rule (from the guide, both products agree on this part): to be tagged, the
page number must be **alone in its own paragraph** -- not embedded within,
or at the start/end of, a paragraph of running text. We tag:
    - plain Arabic numbers:      12          -> $pg12
    - Arabic + letter combos:    A12, 12C, A12B -> $pgA12, $pg12C, $pgA12B
    - continuation ranges:       13-15       -> $pg13-15 (see note below)
    - Roman numerals I..CCC (1..300), upper or lower case

For continuation ranges, the original inserts a DBT-specific continuation
marker meaningful to the Duxbury Braille Translator (written in the source
document as double-bracketed asterisk-delimited codes). That marker has no
meaning for large-print/ODT output, so this filter tags the range as
"$pg13-15" only, without the DBT code. (The Dx_/Braille-side port, if you
build it, is the right place to reproduce that marker exactly --
see docs/COMMAND_MAP.md.)

This filter only *tags*; it does not style the paragraph. Run
lp-format-tagged-page-numbers.lua afterward to turn a tagged paragraph into
the "Print Pg Num" style used by the large print template, and remove the
tag text at that point (matching "Format Tagged Pg Numbers", which removes
the $pg tag once it applies the style).
]]

local ROMAN_PATTERN =
    "^[MCDXLVI]+$"
local roman_lower_pattern = "^[mcdxlvi]+$"

local function is_roman_numeral(s)
    -- Accept only up to CCC (300) per the guide's stated range; a full
    -- roman-to-int validator is out of scope, so just sanity check the
    -- charset and a conservative length limit (CCC = 3 chars, but e.g.
    -- "CCXLVIII" = 248 is 8 chars) -- cap at 8 to avoid matching long
    -- strings of stray letters as if they were numerals.
    if #s < 1 or #s > 8 then
        return false
    end
    return s:match(ROMAN_PATTERN) ~= nil or s:match(roman_lower_pattern) ~= nil
end

-- Matches a paragraph whose ENTIRE stringified content is a reference page
-- number candidate: optional leading letter(s), digits, optional trailing
-- letter(s); or a range "12-15" (with optional leading/trailing letters on
-- either side); or a roman numeral.
local function classify_page_number(text)
    text = text:gsub("^%s+", ""):gsub("%s+$", "")

    -- range: e.g. "13-15", "A19-A21", "22B-26B"
    local a, b = text:match("^(%a?%d+%a?)%-(%a?%d+%a?)$")
    if a and b then
        return "range", text
    end

    -- plain / lettered arabic: "12", "A12", "12C", "A12B"
    if text:match("^%a?%d+%a?$") then
        return "arabic", text
    end

    if is_roman_numeral(text) then
        return "roman", text
    end

    return nil
end

function Para(el)
    local text = pandoc.utils.stringify(el)
    local kind, value = classify_page_number(text)
    if kind then
        -- Replace the whole paragraph's content with the $pg-tagged form.
        -- We keep it as plain text (rather than a Span with metadata) so
        -- lp-format-tagged-page-numbers.lua, or a human editing the
        -- intermediate file, can find it with a plain "$pg" search --
        -- exactly like the Word/LO macros' own approach.
        return pandoc.Para({ pandoc.Str("$pg" .. value) })
    end
    return el
end

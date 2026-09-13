--[[
lp-horizontal-list-to-vertical.lua

Port of "Horizontal List to Vertical" (VistaType LP User Guide p.50-51; same
macro name and behavior in the BANA/Braille guide p.22-23).

Original behavior (interactive): the user places the cursor in -- or selects
-- one paragraph containing answer choices crammed onto one line to save
paper, e.g.:

    a. the Sparrow b. the Fly c. the Fish d. the Beetle

then picks a list type from a dialog (Ordered / Spaced / Tabbed) and the
macro splits it into one paragraph per item:

    a. the Sparrow
    b. the Fly
    c. the Fish
    d. the Beetle

SCOPE FOR THIS BATCH FILTER
============================
Unlike the interactive macro, a Pandoc filter has no user to ask "which
paragraph, which list type?" -- it runs over the whole document. Blindly
splitting on *every* space, or on every tab, would wreck ordinary prose. So
this filter only auto-detects and converts the "Ordered List" case: a
paragraph containing two or more sequential markers of the same family
(letters a./b./c./... , or numbers 1)/2)/3)/...). That pattern essentially
never occurs by accident in ordinary prose, so it's safe to run over an
entire document unattended.

"Spaced List" (bare words separated by run of spaces/tabs with no letter
markers) and "Tabbed List" modes require knowing *which* paragraph the user
means, exactly like the original -- those remain interactive-only and are
implemented in lo-macros/vistatype/lp/formatting_tools.py, operating on the
current selection, matching the original UX.

Sorting: the original defaults to "Sort in Ascending Order" (checkbox
checked). Since markers found by this filter are already, by construction,
in the order they appeared on the line, re-sorting them ascending by marker
is a no-op for well-formed input and only matters for accidentally-scrambled
source lines -- so we do sort by marker value, matching the default.
]]

-- Recognize a token that IS (or starts with) an ordered-list marker:
-- one or two letters, or one or two digits, immediately followed by '.' or
-- ')'. Handles both "b." as a whole Str token and "c.\194\160the" (period +
-- non-breaking space fused to the following word, which Pandoc's own
-- markdown reader produces to avoid list-continuation ambiguity).
local NBSP = "\194\160"

local function match_marker(text)
    local marker, rest = text:match("^(%a%a?[%.%)])(.*)$")
    if not marker then
        marker, rest = text:match("^(%d%d?[%.%)])(.*)$")
    end
    if not marker then
        return nil
    end
    -- strip a fused leading nbsp/space from the remainder, if any
    rest = rest:gsub("^" .. NBSP, ""):gsub("^ +", "")
    return marker, rest
end

local function marker_sort_key(marker)
    local core = marker:sub(1, -2) -- drop trailing punctuation
    local n = tonumber(core)
    if n then
        return n
    end
    return string.byte(core:lower(), 1) - string.byte("a", 1) + 1
end

-- Split a flat Inlines list into segments, each segment starting at a
-- recognized marker. Returns nil if fewer than 3 markers are found, OR if
-- there is any non-whitespace content BEFORE the first marker is found --
-- the latter matters because Pandoc's own Markdown reader sometimes
-- consumes a leading "a." as an *implicit ordered-list start marker*
-- before this filter ever sees the paragraph, leaving the remaining
-- inlines starting mid-item (e.g. "the Sparrow b. the Fly c. ..." with no
-- leading "a."). Converting that case anyway would silently DISCARD "the
-- Sparrow" -- verified this actually happens with a real markdown fixture
-- (pandoc/test/sample.md) before this check was added. Refusing and
-- leaving the block untouched is far safer than guessing.
local function split_on_markers(inlines)
    local segments = {}
    local current_marker = nil
    local current = pandoc.List()
    local saw_content_before_first_marker = false

    local function flush()
        if current_marker then
            table.insert(segments, { marker = current_marker, inlines = current })
        end
    end

    local i = 1
    local n = #inlines
    while i <= n do
        local inline = inlines[i]
        local at_boundary = (i == 1) or (inlines[i - 1].t == "Space")
        if at_boundary and inline.t == "Str" then
            local marker, rest = match_marker(inline.text)
            if marker then
                flush()
                current_marker = marker
                current = pandoc.List()
                if rest ~= "" then
                    current:insert(pandoc.Str(rest))
                end
                i = i + 1
                goto continue
            end
        end
        if current_marker then
            current:insert(inline)
        elseif inline.t ~= "Space" and inline.t ~= "SoftBreak" then
            saw_content_before_first_marker = true
        end
        i = i + 1
        ::continue::
    end
    flush()

    if saw_content_before_first_marker then
        return nil
    end

    -- Require >= 3 markers, not 2. Two sequential markers ("a) the store
    -- and b) the bank") occur often enough in ordinary prose to be an
    -- unsafe trigger for an unattended batch filter; three or more
    -- same-family sequential markers on one line is a pattern that's
    -- specific enough to real horizontal exercise lists to auto-convert
    -- with confidence. (Tested against false-positive prose during
    -- development -- see pandoc/test/horiz2.docx.) If your source has
    -- genuine two-choice horizontal lists that don't get caught, use the
    -- interactive LibreOffice macro instead, which -- like the original --
    -- operates on a paragraph you've selected, not the whole document.
    if #segments < 3 then
        return nil
    end
    return segments
end

local function trim_edges(inlines)
    while #inlines > 0 and (inlines[1].t == "Space" or inlines[1].t == "SoftBreak") do
        inlines:remove(1)
    end
    while #inlines > 0 and (inlines[#inlines].t == "Space" or inlines[#inlines].t == "SoftBreak") do
        inlines:remove(#inlines)
    end
    return inlines
end

local function build_vertical_paras(segments, sort_ascending)
    if sort_ascending ~= false then
        table.sort(segments, function(a, b)
            return marker_sort_key(a.marker) < marker_sort_key(b.marker)
        end)
    end
    local blocks = pandoc.List()
    for _, seg in ipairs(segments) do
        local inlines = pandoc.List()
        inlines:insert(pandoc.Str(seg.marker))
        inlines:insert(pandoc.Space())
        for _, inl in ipairs(trim_edges(seg.inlines)) do
            inlines:insert(inl)
        end
        blocks:insert(pandoc.Para(inlines))
    end
    return blocks
end

local function try_convert_block(block)
    if block.t ~= "Para" and block.t ~= "Plain" then
        return nil
    end
    local segments = split_on_markers(pandoc.List(block.content))
    if not segments then
        return nil
    end
    return build_vertical_paras(segments, true)
end

-- Reconstruct the marker text implied by an OrderedList's own numbering
-- attributes (start/style/delimiter), e.g. start=1, style=LowerAlpha,
-- delimiter=Period -> "a.". Returns nil for a style/delimiter combination
-- we don't recognize, rather than guessing.
local function implicit_marker_text(list_el)
    local n = list_el.start
    local core
    if list_el.style == "LowerAlpha" then
        core = string.char(string.byte("a") + n - 1)
    elseif list_el.style == "UpperAlpha" then
        core = string.char(string.byte("A") + n - 1)
    elseif list_el.style == "Decimal" then
        core = tostring(n)
    else
        return nil
    end
    local punct
    if list_el.delimiter == "OneParen" or list_el.delimiter == "TwoParens" then
        punct = ")"
    elseif list_el.delimiter == "Period" then
        punct = "."
    else
        return nil
    end
    return core .. punct
end

-- Handles the case where Pandoc's own Markdown reader has already
-- consumed a leading "a." (or "1.") as this list's start marker, wrapping
-- the REST of a horizontal list ("the Sparrow b. the Fly c. ...") as the
-- single item's content. Confirmed this actually happens with plain
-- Markdown input (not just a hypothetical): pandoc/test/sample.md's line
-- "a. the Sparrow b. the Fly c. the Fish d. the Beetle" parses as a
-- 1-item OrderedList this way, and naively only handling Para/Plain
-- blocks silently dropped "the Sparrow" (the text between the consumed
-- "a." marker and the next recognized "b." marker) before this function
-- was added -- see git history / docs/FINDINGS.md if this repo tracks
-- that, or just: don't remove this without re-testing sample.md.
function OrderedList(el)
    if #el.content ~= 1 then
        return nil
    end
    local item = el.content[1]
    if #item ~= 1 then
        return nil
    end
    local block = item[1]
    if block.t ~= "Para" and block.t ~= "Plain" then
        return nil
    end
    local marker = implicit_marker_text(el)
    if not marker then
        return nil
    end

    local inlines = pandoc.List()
    inlines:insert(pandoc.Str(marker))
    inlines:insert(pandoc.Space())
    for _, inl in ipairs(block.content) do
        inlines:insert(inl)
    end

    local segments = split_on_markers(inlines)
    if not segments then
        return nil
    end
    return build_vertical_paras(segments, true)
end

function Blocks(blocks)
    local out = pandoc.List()
    for _, block in ipairs(blocks) do
        local converted = try_convert_block(block)
        if converted then
            for _, b in ipairs(converted) do
                out:insert(b)
            end
        else
            out:insert(block)
        end
    end
    return out
end

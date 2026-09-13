--[[
lp-delete-multiple-paragraph-marks.lua

Port of the VistaType LP "Delete Multiple Consecutive Paragraph Marks" macro
(VistaType LP User Guide, "File Cleanup Group", p.45) and the equivalent BANA
macro of the same name.

Word/LP behavior: multiple consecutive empty paragraphs collapse to one.

Pandoc's document model does not represent an "empty paragraph mark" the way
Word does -- a blank line in Markdown is just the block separator and never
becomes a Para at all, and most non-Word sources won't produce empty Para
blocks either. This filter exists mainly for DOCX/ODT input, where a truly
empty paragraph (no runs, or only whitespace/line-break runs) *does* survive
as an empty Para block in the AST. We walk the block list at the document
(and, for completeness, BlockQuote/Div) level and drop consecutive empty
Paras down to a single one.

Usage:
    pandoc input.docx -o output.odt \
        --reference-doc=reference-doc/LargePrintTemplate-reference.odt \
        --lua-filter=pandoc/filters/lp-delete-multiple-paragraph-marks.lua
]]

local NBSP = "\194\160" -- U+00A0 non-breaking space, as UTF-8 bytes

local function is_blank_str(text)
    -- treat a run of only spaces/nbsp as "blank" too -- this is how an
    -- intentionally-empty paragraph sometimes survives round-tripping
    -- through Markdown or ODT (Word itself drops truly-empty paragraphs
    -- from the DOCX reader before they ever reach the AST -- see
    -- docs/COMMAND_MAP.md for the measured behavior).
    return text:gsub(NBSP, " "):match("^%s*$") ~= nil
end

local function is_empty_para(block)
    if block.t ~= "Para" and block.t ~= "Plain" then
        return false
    end
    for _, inline in ipairs(block.content) do
        if inline.t == "Space" or inline.t == "SoftBreak" or inline.t == "LineBreak" then
            -- fine, keeps looking
        elseif inline.t == "Str" and is_blank_str(inline.text) then
            -- fine, keeps looking
        else
            return false
        end
    end
    return true
end

local function collapse(blocks)
    local out = pandoc.List()
    local prev_was_empty = false
    for _, block in ipairs(blocks) do
        local empty = is_empty_para(block)
        if empty and prev_was_empty then
            -- skip: this is the 2nd+ consecutive empty paragraph
        else
            out:insert(block)
        end
        prev_was_empty = empty
    end
    return out
end

function Pandoc(doc)
    doc.blocks = collapse(doc.blocks)
    return doc
end

function BlockQuote(el)
    el.content = collapse(el.content)
    return el
end

function Div(el)
    el.content = collapse(el.content)
    return el
end

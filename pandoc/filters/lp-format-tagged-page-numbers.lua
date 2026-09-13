--[[
lp-format-tagged-page-numbers.lua

Port of "Format Tagged Pg Numbers" (VistaType LP User Guide p.48-49): finds
paragraphs tagged with "$pg" (see lp-tag-reference-page-numbers.lua, or
tags added by hand / by the interactive "Manual Tag Pg Number" LO macro),
applies the template's "Print Pg Num" paragraph style to them, and removes
the "$pg" tag text itself, leaving just the bare reference page number.

This is implemented as a Div wrapper with a `custom-style` attribute, which
Pandoc's ODT writer maps directly onto the named style in the reference
document -- verified against a real (headless) LibreOffice instance: the
resulting paragraph resolves to the actual "Print Pg Num" style from
LargePrintTemplate.dotx, background colour F4A3D4 and all, not just a
same-named placeholder. See lo-macros/tests/check_customstyle.py.

Run this AFTER lp-tag-reference-page-numbers.lua.

Note on continuation ranges ("$pg13-15"): the original macro also encodes a
DBT continuation marker for Braille output, which doesn't apply here; a
large-print reference page range is just left as plain text ("13-15") once
styled.
]]

function Para(el)
    local text = pandoc.utils.stringify(el)
    local value = text:match("^%$pg(.+)$")
    if not value then
        return el
    end
    local styled_para = pandoc.Para({ pandoc.Str(value) })
    return pandoc.Div({ styled_para }, pandoc.Attr("", {}, { { "custom-style", "Print Pg Num" } }))
end

# Abstract base for the per-@odata.type custom documentation handlers populated
# in phase 4. Subclasses live under Internal/Documentation/PolicyTypeHandlers/, declare
# the @odata.type(s) they claim, and override Document() to fill the context.
#
# Registration is by composition (handler instance registers itself with the
# [DocumentationRegistry] at file load time), not by reflection — explicit so
# dispatch order is deterministic and missing registrations are obvious.

class DocumentationHandlerBase {
    [string[]] $ODataTypes

    DocumentationHandlerBase() {
        $this.ODataTypes = @()
    }

    [void] Document([object]$PolicyObject, [DocumentationContext]$Context) {
        throw "DocumentationHandlerBase.Document is abstract - override in derived class $($this.GetType().Name)"
    }
}

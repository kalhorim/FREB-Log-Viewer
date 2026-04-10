#Requires -Version 5.1
<#
.SYNOPSIS
    XML-to-HTML transformation service for the FREB log viewer.

.DESCRIPTION
    Provides Convert-XmlWithXsl, which applies an XSL stylesheet to a FREB XML
    file and returns the resulting HTML as a string.
    All disposable resources (XmlReader, StringWriter) are released in a finally
    block so that resource leaks cannot occur even when a transform error is raised.
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Transforms a FREB XML file with an XSL stylesheet and returns HTML.

.PARAMETER XmlPath
    Absolute path to the FREB XML log file.

.PARAMETER XslPath
    Absolute path to the freb.xsl stylesheet.

.OUTPUTS
    [string]  HTML content ready to be assigned to WebBrowser.DocumentText.
              On failure, returns a minimal HTML page containing the error message.
#>
function Convert-XmlWithXsl {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string] $XmlPath,

        [Parameter(Mandatory)]
        [string] $XslPath
    )

    $xmlReader    = $null
    $stringWriter = $null

    try {
        $xslTransform = [System.Xml.Xsl.XslCompiledTransform]::new()
        $xslTransform.Load($XslPath)

        $xmlReaderSettings                = [System.Xml.XmlReaderSettings]::new()
        $xmlReaderSettings.DtdProcessing  = [System.Xml.DtdProcessing]::Ignore
        $xmlReaderSettings.XmlResolver    = $null

        $stringWriter = [System.IO.StringWriter]::new()
        $xmlReader    = [System.Xml.XmlReader]::Create($XmlPath, $xmlReaderSettings)

        $xslTransform.Transform($xmlReader, $null, $stringWriter)

        return $stringWriter.ToString()
    }
    catch {
        $safeMessage = [System.Web.HttpUtility]::HtmlEncode($_.Exception.Message)
        return "<html><body><pre>Error transforming XML: $safeMessage</pre></body></html>"
    }
    finally {
        if ($null -ne $xmlReader)    { $xmlReader.Dispose() }
        if ($null -ne $stringWriter) { $stringWriter.Dispose() }
    }
}

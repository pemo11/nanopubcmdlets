<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet
    version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt"
    exclude-result-prefixes="msxsl"
  >
  
  <xsl:output
      method="html"
      indent="yes"
      encoding="ISO-8859-1"
  />

  <xsl:template match ="topic">
      <div class="topic">
          <xsl:text>Lektion: </xsl:text><xsl:value-of select="." />
      </div>
  </xsl:template>
  
  <xsl:template match ="subtopic">
      <div class="subtopic">
          <xsl:value-of select="." />
      </div>
  </xsl:template>
  
  <xsl:template match ="para">
      <div class="para">
          <xsl:value-of select="." />
      </div>
  </xsl:template>
  
  <xsl:template match ="/">
      <html>
        <head>
            <title><xsl:value-of select="nanobook/@title" />  </title>
        </head>
        <body>
            <xsl:apply-templates />
        </body>
    </html>
  </xsl:template> 

</xsl:stylesheet>

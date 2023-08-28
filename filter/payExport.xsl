<?xml version="1.0" encoding="UTF-8"?>
<!-- We must define several namespaces, because we need them to access -->
<!-- the document model of the in-memory OpenOffice.org document.      -->
<!-- If we want to access more parts of the document model, we must    -->
<!-- add there namesspaces here, too.                                  -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
   xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
   xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
   xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
   exclude-result-prefixes="office table text">

<xsl:output method = "xml" indent = "yes" encoding = "UTF-8" omit-xml-declaration = "no"/>

<!-- By setting the PropertyValue "URL" in the properties used in storeToURL(), -->
<!-- we can pass a single parameter to this stylesheet.                         -->
<!-- Caveat: If we use the "URL" property in the stylesheet and call in OOo     -->
<!-- from the menu "File" > "Export...", OOo assigns a target URL. And that     -->
<!-- might not be what we want.                                                 -->
<xsl:param name="targetURL"/>

<xsl:variable name="exportDate">
  <xsl:choose>
   <xsl:when test="string-length(substring-before($targetURL,';'))=10">
    <xsl:value-of select="substring-before($targetURL,';')"/>
   </xsl:when>
   <xsl:when test="string-length($targetURL)=10">
    <xsl:value-of select="$targetURL"/>
   </xsl:when>
  </xsl:choose>
</xsl:variable>

<xsl:variable name="exportUser">
  <xsl:if test="string-length(substring-after($targetURL,';'))>0">
   <xsl:value-of select="substring-after($targetURL,';')"/>
  </xsl:if>
</xsl:variable>

<!-- Process the document model -->
<xsl:template match="/">
  <payments>
   <xsl:attribute name="export-date"><xsl:value-of select="$exportDate"/></xsl:attribute>
   <xsl:attribute name="export-user"><xsl:value-of select="$exportUser"/></xsl:attribute>
   <!-- Process all tables -->
   <xsl:apply-templates select="//table:table"/>
  </payments>
</xsl:template>

<xsl:template match="table:table">
  <!-- Process all table-rows after the column labels in table-row 1 -->
  <xsl:for-each select="table:table-row">
   <xsl:if test="position()>1">
    <payment>
     <!-- Process the first for columns containing purpose, amount, tax and maturity -->
     <xsl:for-each select="table:table-cell">
      <xsl:choose>
       <xsl:when test="position()=1">
       <purpose><xsl:value-of select="text:p"/></purpose>
       </xsl:when>
       <xsl:when test="position()=2">
       <amount><xsl:value-of select="@office:value"/></amount>
      </xsl:when>
       <xsl:when test="position()=3">
       <tax><xsl:value-of select="@office:value"/></tax>
      </xsl:when>
       <xsl:when test="position()=4">
       <maturity><xsl:value-of select="@office:date-value"/></maturity>
      </xsl:when>
      </xsl:choose>
     </xsl:for-each>
    </payment>
   </xsl:if>
  </xsl:for-each>
</xsl:template>

</xsl:stylesheet>
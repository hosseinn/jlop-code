<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  exclude-result-prefixes="fo office style text table draw xlink number form script svg">

<xsl:output method="xml" indent="yes" encoding = "UTF-8" omit-xml-declaration = "no"/>

<xsl:template match="/">
  <club-database> 
    <xsl:apply-templates select="//office:text/text:h[@text:style-name='Association']"/>
  </club-database>
</xsl:template>

<xsl:template match="text:h[@text:style-name='Association']">
    <association id="{.}">
        <xsl:call-template name="make-club">
            <xsl:with-param name="clubNode"
             select="following-sibling::text:h[1]"/>
        </xsl:call-template>
    </association>
</xsl:template>

<xsl:template name="make-club">
    <xsl:param name="clubNode"/>
    <xsl:if test="$clubNode/@text:style-name = 'Club_20_Name'">
        <club>
            <xsl:attribute name="id">
                <xsl:value-of
                    select="$clubNode/text:span[@text:style-name='Club_20_Code']"/>
            </xsl:attribute>        
            <name><xsl:value-of select="$clubNode"/></name>
            <xsl:call-template name="make-content">
                <xsl:with-param name="contentNode"
                    select="$clubNode/following-sibling::*[1]"/>
            </xsl:call-template>
            
        </club>
        <xsl:if test="$clubNode/following-sibling::text:h[1]">
            <xsl:call-template name="make-club">
                <xsl:with-param name="clubNode"
                    select="$clubNode/following-sibling::text:h[1]"/>
            </xsl:call-template>
        </xsl:if>
    </xsl:if>
</xsl:template>

<xsl:template name="make-content">
    <xsl:param name="contentNode"/>
    <xsl:if test="name($contentNode) = 'text:p'">
        <xsl:choose>
            <xsl:when test="$contentNode/text:span">
                <xsl:call-template name="add-item">
                    <xsl:with-param name="spanNode"
                        select="$contentNode/text:span"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:when test="name($contentNode/following-sibling::*[1]) = 'text:list'">
                <xsl:call-template name="email-list">
                    <xsl:with-param name="emailList"
                        select="$contentNode/following-sibling::text:list[1]"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:when test="$contentNode/@text:style-name = 'Club_20_Info'">
                <info>
                    <xsl:apply-templates select="$contentNode"/>
                </info>
            </xsl:when>
        </xsl:choose>
        <xsl:call-template name="make-content">
            <xsl:with-param name="contentNode"
                select="$contentNode/following-sibling::*[1]"/>
        </xsl:call-template>
    </xsl:if>
</xsl:template>

<xsl:template name="add-item">
    <xsl:param name="spanNode"/>
    <xsl:variable name="styleAttr" select="$spanNode/@text:style-name"/>
    
    <xsl:choose>
        <xsl:when test="$styleAttr = 'Charter'">
            <charter><xsl:value-of select="$spanNode"/></charter>
        </xsl:when>
        <xsl:when test="$styleAttr = 'Contact'">
            <contact><xsl:value-of select="$spanNode"/></contact>
        </xsl:when>
        <xsl:when test="$styleAttr = 'Phone'">
            <phone><xsl:value-of select="$spanNode"/></phone>
        </xsl:when>
        <xsl:when test="$styleAttr = 'Location'">
            <location><xsl:value-of select="$spanNode"/></location>
        </xsl:when>
        <xsl:when test="$styleAttr = 'Email'">
            <email><xsl:value-of select="$spanNode"/></email>
        </xsl:when>
        <xsl:when test="$styleAttr = 'Age_20_Groups'">
            <age-groups>
                <xsl:attribute name="type">
                    <xsl:value-of select="translate($spanNode,' abcdefghijklmnopqrstuvwxyz', '')"/>
                </xsl:attribute>
            </age-groups>
        </xsl:when>
    </xsl:choose>
</xsl:template>

<xsl:template name="email-list">
    <xsl:param name="emailList"/>
    <xsl:for-each select="$emailList/descendant::text:span[@text:style-name='Email']">
        <email><xsl:value-of select="."/></email>
    </xsl:for-each>
</xsl:template>

<xsl:template match="text:a">
<a href="{@xlink:href}"><xsl:apply-templates/></a>
</xsl:template>

</xsl:stylesheet>


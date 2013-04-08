<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    version="1.0">

    <xsl:output method="text" />
    <xsl:strip-space elements="*"/>
    
    <!-- Document that contains the numbering styles for paragraphs -->
    <xsl:variable name="numberingStyles" select="document('numbering.xml')"/>
    
    <!-- Add line break after every paragraph -->
    <!-- TODO: Add additional line break if preceding p is a list -->
    <xsl:template match="w:p">
        <xsl:apply-templates select="w:pPr" mode="paragraph-prefix" />
        <xsl:apply-templates select="w:r"/>
        <xsl:apply-templates select="w:pPr" mode="paragraph-suffix" />
        <xsl:text xml:space="preserve">&#10;</xsl:text>
    </xsl:template>
    
    <!-- Format region -->
    <xsl:template match="w:r">
        <xsl:text xml:space="preserve"> </xsl:text>
        <xsl:apply-templates select="w:rPr" mode="inline-prefix" />
        <xsl:apply-templates select="w:t|w:br"/>
        <xsl:apply-templates select="w:rPr" mode="inline-suffix" />
        <xsl:text xml:space="preserve"> </xsl:text>
    </xsl:template>
    
    
    <!-- Lists -->
    <xsl:template match="w:numPr" mode="paragraph-prefix">
        <xsl:variable name="ilvl" select="w:ilvl/@w:val"></xsl:variable>
        <xsl:variable name="prefix">
            <xsl:call-template name="getListprefix">
                <xsl:with-param name="numId" select="w:numId/@w:val"/>
                <xsl:with-param name="ilvl" select="$ilvl"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$prefix">
            <xsl:call-template name="getSpaces">
                <xsl:with-param name="numSpaces" select="($ilvl+1)*2"></xsl:with-param>
            </xsl:call-template>
            <xsl:value-of select="$prefix"/>
        </xsl:if>
    </xsl:template>
    
    
    <!-- Bold text inline -->
    <xsl:template match="w:b" mode="inline-prefix">**</xsl:template>
    <xsl:template match="w:b" mode="inline-suffix">**</xsl:template>
    
    <!-- Italic text inline -->
    <xsl:template match="w:i" mode="inline-prefix">//</xsl:template>
    <xsl:template match="w:i" mode="inline-suffix">//</xsl:template>
    
    <!-- Add line breaks -->
    <xsl:template match="w:br">
        <xsl:text xml:space="preserve"> \\ </xsl:text>
    </xsl:template>
    
    
    <!-- Check for a matching numbering style and return a prefix if numbering style is decimal or bullet -->
    <xsl:template name="getListprefix">
        <xsl:param name="numId"/>
        <xsl:param name="ilvl"/>
        
        <xsl:variable name="numberElement" 
            select="$numberingStyles/w:numbering/w:abstractNum[@w:abstractNumId=$numId]/w:lvl[@w:ilvl=$ilvl]"/>
        
        <xsl:choose>
            <xsl:when test="$numberElement/w:numFmt/@w:val = 'decimal'">
                <xsl:text xml:space="preserve">- </xsl:text>
            </xsl:when>
            <xsl:when test="$numberElement/w:numFmt/@w:val = 'bullet'">
                <xsl:text xml:space="preserve">* </xsl:text>
            </xsl:when>
        </xsl:choose>
        
    </xsl:template>
    
    <xsl:template name="getSpaces">
        <xsl:param name="numSpaces" select="1"></xsl:param>
        <xsl:value-of select="substring('                    ', 0, $numSpaces)" xml:space="preserve" />
    </xsl:template>
    

</xsl:stylesheet>

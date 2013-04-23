<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    version="1.0">

    <xsl:output method="text" />
    
    
    <xsl:param name="numberingFile" select="'numbering.xml'"></xsl:param>
    <xsl:param name="styleFormattingFile" select="'styleformatting.xml'"></xsl:param>
        
    <!-- 
        Document that contains the numbering styles for paragraphs.
        Normally this is 'numbering.xml' contained in the docx archive file.
    -->
    <xsl:variable name="numberingStyles" select="document($numberingFile)"/>
    
    <!-- 
        Document that contains the associations between style names and other formatting
        Used for headings, blockquotes etc
    -->
    <xsl:variable name="styleFormatting" select="document($styleFormattingFile)/formatting"/>
    
    <!-- Add line break after every paragraph -->
    <!-- TODO: Add additional line break if preceding p is a list -->
    <xsl:template match="w:p">
        <xsl:apply-templates select="w:pPr" mode="paragraph-prefix" />
        <xsl:apply-templates select="w:r|w:hyperlink"/>
        <xsl:apply-templates select="w:pPr" mode="paragraph-suffix" />
        <xsl:text xml:space="preserve">&#10;</xsl:text>
    </xsl:template>
    
    <!-- Format region -->
    <xsl:template match="w:r">
        <xsl:apply-templates select="w:rPr" mode="inline-prefix" />
        <xsl:apply-templates select="w:t|w:br|w:sym"/>
        <xsl:apply-templates select="w:rPr" mode="inline-suffix" />
    </xsl:template>
    
    <!-- Lists -->
    <xsl:template match="w:numPr" mode="paragraph-prefix">
        <xsl:variable name="ilvl" select="number(w:ilvl/@w:val)"></xsl:variable>
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
    
    <!-- Paragraph prefixes based on style -->
    <xsl:template match="w:pStyle" mode="paragraph-prefix">
        <xsl:variable name="stylename" select="@w:val"/>
        <xsl:value-of select="$styleFormatting/style[@name=$stylename]/prefix"/>
    </xsl:template>
    
    <xsl:template match="w:pStyle" mode="paragraph-suffix">
        <xsl:variable name="stylename" select="@w:val"/>
        <xsl:value-of select="$styleFormatting/style[@name=$stylename]/suffix"/>
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
    
    <!-- Table Handling -->
    <xsl:template match="w:tbl">
        <xsl:apply-templates select="w:tr"/>
        <xsl:text>&#10;</xsl:text>
    </xsl:template>
    
    <!-- Table column -->
    <xsl:template match="w:tc">
        <xsl:if test="count(preceding-sibling::w:tc) = 0">
            <xsl:text>&#10;|</xsl:text>
        </xsl:if>
        <xsl:value-of select="w:p"/>
        <xsl:text>|</xsl:text>
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
        <xsl:value-of select="substring('                    ', 1, $numSpaces)" xml:space="preserve" />
    </xsl:template>
    
    <!-- Symbols -->
    <xsl:template match="w:sym[@w:font='Wingdings' and @w:char='F0E0']">
        <xsl:text> =></xsl:text>
    </xsl:template>

</xsl:stylesheet>

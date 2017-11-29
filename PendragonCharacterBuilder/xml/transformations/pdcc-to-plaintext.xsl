<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    exclude-result-prefixes="xs"
    version="1.0">
    
    <xsl:output indent="no" omit-xml-declaration="yes" method="text"/>
    <xsl:strip-space elements="*"/>
    
    <!-- Have I mentioned how much I hate that Microsoft won't support
        XSL 2.0? As someone who works with XML and VB.NET I hate it...
        quite a lot. -->
    <xsl:variable name="lowercase" select="'abcdefghijklmnopqrstuvwxyz'" />
    <xsl:variable name="uppercase" select="'ABCDEFGHIJKLMNOPQRSTUVWXYZ'" />
    <xsl:variable name="nl"><xsl:text>&#10;</xsl:text></xsl:variable>
    
    <xsl:template match="/">
        <xsl:apply-templates/>
    </xsl:template>
    
    <xsl:template match="character">
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">1</xsl:with-param>
            <xsl:with-param name="t">
                <xsl:value-of select="concat(
                    translate(substring(personal-data/current-class, 1, 1), 
                    $lowercase, $uppercase),
                    substring(personal-data/current-class, 2), ' ',
                    personal-data/name,
                    $nl
                    )"/>
            </xsl:with-param>
        </xsl:call-template>
        
        <xsl:call-template name="stat-to-line">
            <xsl:with-param name="elem" select="personal-data/age"/>
        </xsl:call-template>
        <xsl:call-template name="stat-to-line">
            <xsl:with-param name="elem" select="personal-data/gender"/>
        </xsl:call-template>
        <xsl:call-template name="statline-with-colon">
            <xsl:with-param name="pre-colon">Child Number</xsl:with-param>
            <xsl:with-param name="post-colon">
                <xsl:value-of select="personal-data/child-number"/>
            </xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="stat-to-line">
            <xsl:with-param name="elem" select="personal-data/glory"/>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        
        <xsl:call-template name="stat-to-line">
            <xsl:with-param name="elem" select="personal-data/culture"/>
        </xsl:call-template>
        <xsl:call-template name="stat-to-line">
            <xsl:with-param name="elem" select="personal-data/homeland"/>
        </xsl:call-template>
        <xsl:call-template name="statline-with-colon">
            <xsl:with-param name="pre-colon">Home</xsl:with-param>
            <xsl:with-param name="post-colon">
                <xsl:value-of select="personal-data/current-home"/>
            </xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="stat-to-line">
            <xsl:with-param name="elem" select="personal-data/lord"/>
        </xsl:call-template>
        <xsl:call-template name="stat-to-line">
            <xsl:with-param name="elem" select="personal-data/religion"/>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">2</xsl:with-param>
            <xsl:with-param name="t">Attributes</xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        
        <xsl:for-each select="attributes/base-attributes/*">
            <xsl:call-template name="attribute-output">
                <xsl:with-param name="elem" select="."/>
            </xsl:call-template>
        </xsl:for-each>
        <xsl:text>&#10;</xsl:text>
        
        <xsl:for-each select="attributes/derived-attributes/*">
            <xsl:call-template name="stat-to-line">
                <xsl:with-param name="elem" select="."/>
            </xsl:call-template>
        </xsl:for-each>
        <xsl:text>&#10;</xsl:text> 
        
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">3</xsl:with-param>
            <xsl:with-param name="t">Distinctive Features</xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        <xsl:value-of select="features"/>
        <xsl:text>&#10;</xsl:text>
        <xsl:text>&#10;</xsl:text>
        
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">2</xsl:with-param>
            <xsl:with-param name="t">Traits</xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        <xsl:for-each select="personality-traits/trait-pair">
            <xsl:call-template name="trait-output">
                <xsl:with-param name="pair" select="."/>
            </xsl:call-template>
        </xsl:for-each>
        <xsl:text>&#10;</xsl:text>
        
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">2</xsl:with-param>
            <xsl:with-param name="t">Passions</xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        <xsl:for-each select="passions/passion">
            <xsl:call-template name="passion-output">
                <xsl:with-param name="passion" select="."/>
            </xsl:call-template>
        </xsl:for-each>
        <xsl:text>&#10;</xsl:text>
        
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">2</xsl:with-param>
            <xsl:with-param name="t">Skills</xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">3</xsl:with-param>
            <xsl:with-param name="t">Non-Combat</xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        <xsl:for-each select="skills/non-combat/skill">
            <xsl:if test="./@value > 0">
                <xsl:call-template name="skill-output">
                    <xsl:with-param name="skill" select="."/>
                </xsl:call-template>
            </xsl:if>
        </xsl:for-each>
        <xsl:text>&#10;</xsl:text>
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">3</xsl:with-param>
            <xsl:with-param name="t">Combat</xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        <xsl:for-each select="skills/combat/skill">
            <xsl:if test="./@value > 0">
                <xsl:call-template name="skill-output">
                    <xsl:with-param name="skill" select="."/>
                </xsl:call-template>
            </xsl:if>
        </xsl:for-each>
        <xsl:text>&#10;</xsl:text>
        
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">2</xsl:with-param>
            <xsl:with-param name="t">
                <xsl:choose>
                    <xsl:when test="personal-data/current-class = 'lady'">Lady in Waiting</xsl:when>
                    <xsl:otherwise>Squire</xsl:otherwise>
                </xsl:choose>
            </xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        <xsl:call-template name="stat-to-line">
            <xsl:with-param name="elem" select="squire/name"/>
        </xsl:call-template>
        <xsl:call-template name="stat-to-line">
            <xsl:with-param name="elem" select="squire/age"/>
        </xsl:call-template>
        <xsl:for-each select="squire/skill">
            <xsl:if test="./@value > 0">
                <xsl:call-template name="skill-output">
                    <xsl:with-param name="skill" select="."/>
                </xsl:call-template>
            </xsl:if>
        </xsl:for-each>
        <xsl:text>&#10;</xsl:text>
        
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">2</xsl:with-param>
            <xsl:with-param name="t">Stuff</xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        <xsl:for-each select="stuff/item">
            <xsl:value-of select="concat(., $nl)"/>
        </xsl:for-each>
        <xsl:text>&#10;</xsl:text>
        
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">2</xsl:with-param>
            <xsl:with-param name="t">Horses</xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        <xsl:for-each select="horses/horse">
            <xsl:value-of select="concat(./@type, $nl)"/>
        </xsl:for-each>
        <xsl:text>&#10;</xsl:text>
    </xsl:template>
    
    <xsl:template match="history">
        <xsl:call-template name="heading">
            <xsl:with-param name="h-level">2</xsl:with-param>
            <xsl:with-param name="t">History</xsl:with-param>
        </xsl:call-template>
        <xsl:text>&#10;</xsl:text>
        <xsl:for-each select="./p">
            <xsl:value-of select="concat(., $nl, $nl)"/>
        </xsl:for-each>
    </xsl:template>
    
    <xsl:template match="family">
        
    </xsl:template>
    
    <xsl:template name="heading">
        <xsl:param name="h-level"/>
        <xsl:param name="t"/>
        <xsl:choose>
            <xsl:when test="$h-level = 1">
                <xsl:value-of select="translate($t, $lowercase, $uppercase)"/>
            </xsl:when>
            <xsl:when test="$h-level = 2">
                <xsl:value-of select="translate($t, $lowercase, $uppercase)"/>
            </xsl:when>
            <xsl:when test="$h-level = 3">
                <xsl:value-of select="$t"/>
            </xsl:when>
            <xsl:otherwise>
                <!-- A failsafe here. -->
                <xsl:value-of select="$t"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <xsl:template name="stat-to-line">
        <xsl:param name="elem"/>
        <xsl:variable name="a" select="local-name($elem)"/>
        <xsl:variable name="b" select="$elem/text()"/>
        <xsl:call-template name="statline-with-colon">
            <xsl:with-param name="pre-colon">
                <xsl:value-of select="$a"/>
            </xsl:with-param>
            <xsl:with-param name="post-colon">
                <xsl:value-of select="$b"/>
            </xsl:with-param>
        </xsl:call-template>
    </xsl:template>
    
    <xsl:template name="statline-with-colon">
        <xsl:param name="pre-colon"/>
        <xsl:param name="post-colon"/>
        <xsl:value-of select="concat(
            translate(substring($pre-colon, 1, 1), $lowercase, $uppercase),
            substring($pre-colon, 2), ': ',
            translate(substring($post-colon, 1, 1), $lowercase, $uppercase),
            substring($post-colon, 2), $nl
            )"/>
    </xsl:template>

    <xsl:template name="attribute-output">
        <xsl:param name="elem"/>
        <xsl:value-of select="concat(
            $elem/@short,
            ' ',
            $elem/text(),
            $nl
            )"/>
    </xsl:template>

    <xsl:template name="trait-output">
        <xsl:param name="pair"/>
        <xsl:variable name="t1">
            <xsl:value-of select="substring(
                concat($pair/trait[1]/@name, '            '), 
                1, 12)"/>
            <!-- This is such a clever operation! Thanks Stack Overflow! -->
        </xsl:variable>
        <xsl:variable name="t2">
            <xsl:value-of select="concat(
                substring('            ',
                    string-length($pair/trait[2]/@name)),
                $pair/trait[2]/@name
                )"/>
        </xsl:variable>
        <xsl:variable name="v1">
            <xsl:choose>
                <xsl:when test="$pair/trait[1]/@value >= 10">
                    <xsl:value-of select="$pair/trait[1]/@value"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="concat(' ', $pair/trait[1]/@value)"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="v2">
            <xsl:value-of select="substring(
                concat(
                $pair/trait[2]/@value, 
                '  '), 
                1, 2
                )"/>
        </xsl:variable>
        <xsl:value-of select="concat(
            $t1, 
            $v1,
            ' / ',
            $v2,
            $t2,
            $nl
            )"/>
    </xsl:template>

    <xsl:template name="passion-output">
        <xsl:param name="passion"/>
        <xsl:variable name="p">
            <xsl:value-of select="substring(concat($passion/@name, '                  '), 1, 17)"/>
        </xsl:variable>
        <xsl:variable name="v">
            <xsl:choose>
                <xsl:when test="$passion/@value >= 10">
                    <xsl:value-of select="$passion/@value"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="concat(' ', $passion/@value)"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:value-of select="concat($p, $v, $nl)"/>
    </xsl:template>

    <xsl:template name="skill-output">
        <xsl:param name="skill"/>
        <xsl:variable name="s">
            <xsl:value-of select="substring(concat($skill/@name, 
                '                                   '), 1, 32)"/>
        </xsl:variable>
        <xsl:variable name="v">
            <xsl:choose>
                <xsl:when test="$skill/@value >= 10">
                    <xsl:value-of select="$skill/@value"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="concat(' ', $skill/@value)"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:value-of select="concat($s, $v, $nl)"/>
    </xsl:template>
</xsl:stylesheet>
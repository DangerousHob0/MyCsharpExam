<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">

<xsl:template match="/">
<xsl:text disable-output-escaping="yes">&lt;xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:aml="http://schemas.microsoft.com/aml/2001/core" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wsp="http://schemas.microsoft.com/office/word/2003/wordml/sp2" xmlns:sl="http://schemas.microsoft.com/schemaLibrary/2003/core" w:macrosPresent="no" w:embeddedObjPresent="no" w:ocxPresent="no" xml:space="preserve"&gt;</xsl:text>
<xsl:apply-templates mode="getParams"/>
<xsl:text disable-output-escaping="yes">&lt;xsl:template match="/"&gt;</xsl:text>
<xsl:text disable-output-escaping="yes">&lt;xsl:processing-instruction name="mso-application"&gt;&lt;xsl:text&gt;progid="Word.Document"&lt;/xsl:text&gt;&lt;/xsl:processing-instruction&gt;</xsl:text>			
<w:wordDocument xmlns:aml="http://schemas.microsoft.com/aml/2001/core" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wsp="http://schemas.microsoft.com/office/word/2003/wordml/sp2" xmlns:sl="http://schemas.microsoft.com/schemaLibrary/2003/core" w:macrosPresent="no" w:embeddedObjPresent="no" w:ocxPresent="no" xml:space="preserve">
	<w:ignoreSubtree w:val="http://schemas.microsoft.com/office/word/2003/wordml/sp2"/>
	<o:DocumentProperties>
				<o:Title/>
				<o:Author/>
				<o:LastAuthor/>
				<o:Revision/>
				<o:TotalTime/>
				<o:LastPrinted/>
				<o:Created/>
				<o:LastSaved/>
				<o:Pages/>
				<o:Words/>
				<o:Characters/>
				<o:Company/>
				<o:Lines/>
				<o:Paragraphs/>
				<o:CharactersWithSpaces/>
				<o:Version/>
			</o:DocumentProperties>
	<xsl:apply-templates select="//w:body" mode="copy"/>
</w:wordDocument>
<xsl:text disable-output-escaping="yes">&lt;/xsl:template&gt;</xsl:text>	
<xsl:text disable-output-escaping="yes">&lt;/xsl:stylesheet&gt;</xsl:text>
</xsl:template>


<!-- Copy -->
<xsl:attribute-set name = "name-list"/>
<xsl:template match="*|@*" mode="copy">
	<xsl:copy use-attribute-sets="name-list">
		<xsl:apply-templates select="*|@*|text()" mode="copy"/>
	</xsl:copy>		
</xsl:template>

<xsl:template match="text()" mode="copy" priority="1">
	<xsl:value-of select="."/>
</xsl:template> 
<xsl:template match="w:t" mode="copy" priority="2">
<w:t>
	<xsl:choose>
		<xsl:when test="../w:rPr/w:color[@w:val = 'FF0000'] and starts-with(.,'{') and substring(.,string-length(.),1)='}'">
		<xsl:text disable-output-escaping="yes">&lt;xsl:value-of select="$</xsl:text>
		<xsl:value-of select="substring(.,2,string-length(.)-2)"/>
		<xsl:text disable-output-escaping="yes">"/&gt;</xsl:text>
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select="."/>
		</xsl:otherwise>
	</xsl:choose>
</w:t>
</xsl:template>
<!-- Copy -->

<!-- GetParams -->
<xsl:template match="text()" mode="getParams"/>

<xsl:key name="uniq_param" match="//w:t[../w:rPr/w:color[@w:val = 'FF0000'] and starts-with(.,'{')]" use="."/>

<xsl:template match="w:t" mode="getParams">
	<xsl:if test ="../w:rPr/w:color[@w:val = 'FF0000'] and starts-with(.,'{') and substring(.,string-length(.),1)='}'">
		<xsl:if test="generate-id(.) = generate-id(key('uniq_param',.))">
			<xsl:text disable-output-escaping="yes">&lt;xsl:param name=&quot;</xsl:text>
			<xsl:value-of select="substring(.,2,string-length(.)-2)"/>		
			<xsl:text disable-output-escaping="yes">&quot;&gt;&lt;/xsl:param&gt;</xsl:text>		
		</xsl:if>
	</xsl:if>
</xsl:template> 
<!-- GetParams -->

</xsl:stylesheet>
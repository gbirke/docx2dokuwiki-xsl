#!/bin/sh

xsltproc \
	--stringparam numberingFile Wiki_test2_docx_expanded/word/numbering.xml \
	--stringparam relationsFile Wiki_test2_docx_expanded/word/_rels/document.xml.rels \
	docx2wiki.xsl Wiki_test2_docx_expanded/word/document.xml

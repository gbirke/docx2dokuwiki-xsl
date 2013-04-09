#!/bin/sh

xsltproc \
	--stringparam numberingFile Wiki_test_docx_expanded/word/numbering.xml \
	docx2wiki.xsl Wiki_test_docx_expanded/word/document.xml

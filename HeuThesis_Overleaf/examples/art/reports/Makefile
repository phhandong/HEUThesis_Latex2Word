# =============================================================
#                 LaTeX Makefile
# =============================================================
# Usage:
#   make           - ÍêÈ«±àÒë
#   make clean     - ÇåÀí¸¨ÖúÎÄ¼þ
#   make cleanall  - ÇåÀíPDF¼°ËùÓÐ¸¨ÖúÎÄ¼þ
#   make build     - ±àÒë²¢ÇåÀí¸¨ÖúÎÄ¼þ
#   make rebuild   - ÖØÐÂ±àÒë²¢ÇåÀí¸¨ÖúÎÄ¼þ
#   make help      - ÏÔÊ¾°ïÖú
# =============================================================

# ¶¨Òå±àÒëÆ÷¼°±àÒë²ÎÊý
TEXCMD = latexmk -xelatex -interaction=nonstopmode -halt-on-error

# ¶¨ÒåÐèÒªÇåÀíµÄ¸¨ÖúÎÄ¼þÀ©Õ¹Ãû
AUX_FILES = *.acn *.acr *.alg *.aux *.bak *.bbl *.bcf *.blg *.brf *.cut *.dvi *.exa *.fen *.fls *.glg *.glo *.gls *.ilg *.ind *.ist *.lof *.log *.lot *.nav *.snm *.out *.thm *.ten *.loa *.toc *.toe *.vrb *.xdv *.fdb_latexmk *.run.xml *.synctex.gz

# ¶¨ÒåÅÅ³ýµÄÄ¿Â¼(¿ÉÑ¡)
EXCLUDE_DIRS = .git/

# »ñÈ¡µ±Ç°Ä¿Â¼ÏÂËùÓÐ .tex ÎÄ¼þ£¨ÅÅ³ý´ø¿Õ¸ñµÄÎÄ¼þÃû£©
MAIN_TEX  = report
MAIN_PDF  = $(MAIN_TEX).pdf
TEX_FILES = $(wildcard *.tex) $(wildcard chapter/*.tex) $(wildcard appendix/*.tex) $(wildcard preface/*.tex)
BIB_FILES = $(wildcard *.bib)

# Ä¬ÈÏÄ¿±ê£º±àÒëËùÓÐ .tex ÎÄ¼þ
all:  $(MAIN_PDF)
	@echo "±àÒëÍê³É£¡"

# Ä£Ê½¹æÔò£ºxeLaTeX ¡ú BibTeX ¡ú xeLaTeX ¡ú xeLaTeX
$(MAIN_PDF): $(TEX_FILES)  $(BIB_FILES)
	@echo "ÕýÔÚ½øÐÐ±àÒë..."
	$(TEXCMD) $(MAIN_TEX)

# ÇåÀíËùÓÐ¸¨ÖúÎÄ¼þ
clean:
	@echo "ÕýÔÚÇåÀí¸¨ÖúÎÄ¼þ..."
	@for pattern in $(AUX_FILES); do \
		#@echo "ÕýÔÚÉ¾³ý $$pattern ÎÄ¼þ:"; \
		find . -type f -name "$$pattern" ! -path "./$(EXCLUDE_DIRS)/*" -print -delete || true; \
	done
	@echo "ÇåÀíÍê³É£¡"

# ÇåÀíËùÓÐÉú³ÉÎÄ¼þ
cleanall:
	@echo "ÕýÔÚÇåÀíÎÄ¼þ..."
	@for pattern in $(AUX_FILES); do \
		#@echo "ÕýÔÚÉ¾³ý $$pattern ÎÄ¼þ:"; \
		find . -type f -name "$$pattern" ! -path "./$(EXCLUDE_DIRS)/*" -print -delete || true; \
	done
	@rm -f $(MAIN_PDF)
	@echo "ÇåÀíÍê³É£¡"
	
# ±àÒë²¢ÇåÀí£¨³£ÓÃÄ¿±ê£©
build: all clean

# ÖØÐÂ±àÒë²¢ÇåÀí
rebuild: cleanall all clean

help:
	@echo "Usage:"
	@echo "make          - ÍêÈ«±àÒë"
	@echo "make clean    - ÇåÀí¸¨ÖúÎÄ¼þ"
	@echo "make rmpdf    - ÇåÀíPDF¼°ËùÓÐ¸¨ÖúÎÄ¼þ"
	@echo "make build    - ±àÒë²¢ÇåÀí¸¨ÖúÎÄ¼þ"
	@echo "make rebuild  - ÖØÐÂ±àÒë²¢ÇåÀí¸¨ÖúÎÄ¼þ"
	@echo "make help     - ÏÔÊ¾°ïÖú"

.PHONY: all clean cleanall build rebuild help

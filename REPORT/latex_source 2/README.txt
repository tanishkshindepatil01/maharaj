Edge-Deployable Rock Detection and Burial Status Classification
Figure Formatting Revision - LaTeX Source

Compile the report from main.tex with pdfLaTeX and BibTeX. A typical local
workflow is:

    latexmk -pdf main.tex

The source uses the supplied University of Bristol dissertation.cls file,
which is included unchanged. Keep dissertation.cls, sample_bibtex.bib,
logo_uob_color.eps (or the supplied PDF equivalent), and the figures folder
beside main.tex when compiling. The algorithm2e package is also required and
is available in standard full TeX Live and Overleaf installations.

The scripts/generate_revision_charts.py script reproduces the replacement
Figure 3.7, the graphical summary of Table 4.1, and the deployment
speed-versus-mAP50--95 scatter plot. It requires Python with Matplotlib and
NumPy. The generated PNG files are already included, so running the script is
not necessary to compile the report.

Included revisions:
- expanded acknowledgements;
- descriptive captions and logically grouped figure placement;
- expanded YOLO detection and output-processing discussion;
- Figure 3.7 presented as three individual images with an independent
  explanation for each panel;
- removal of Figure 3.8;
- Figure 4.1 regrouped as three paired-chart images with its explanation split
  into efficiency/resources and detection-quality parts;
- deployment FPS and mAP50--95 compared in a labelled scatter plot in
  Section 4.3;
- all numbered pre-appendix content retained within the 30-page limit.

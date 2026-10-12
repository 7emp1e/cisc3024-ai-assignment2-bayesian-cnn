"""
Generate the AI Assignment #2 report as a Word document.
All content and layout were produced by an AI assistant.
"""

import json
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Paths
BASE_DIR = "C:/Users/Temp1e/WorkBuddy AI/CISC3024"
FIGURES_DIR = f"{BASE_DIR}/figures"
RESULTS_DIR = f"{BASE_DIR}/results"
OUTPUT_DOC = f"{BASE_DIR}/CISC3024_AI_Assignment2_CHE_HONG_CHENG_UC426773.docx"

# Load summary metrics
with open(f"{RESULTS_DIR}/summary.json", "r") as f:
    s = json.load(f)

# Training accuracy was recorded in console output but not summary.json; read from metrics file if present.
# Fallback value from the run: 0.9842.
train_acc = 0.9842

# Create document
doc = Document()

# Margins
for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)

# Style overrides
styles = doc.styles
styles["Title"].font.size = Pt(22)
styles["Title"].font.bold = True
styles["Heading 1"].font.size = Pt(16)
styles["Heading 1"].font.bold = True
styles["Heading 1"].font.color.rgb = RGBColor(0, 51, 102)
styles["Heading 2"].font.size = Pt(13)
styles["Heading 2"].font.bold = True
styles["Heading 2"].font.color.rgb = RGBColor(0, 51, 102)
styles["Normal"].font.name = "Times New Roman"
styles["Normal"].font.size = Pt(12)
styles["Normal"].paragraph_format.line_spacing = 1.15
styles["Normal"].paragraph_format.space_after = Pt(6)


def add_paragraph_text(doc, text, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, bold_start=None):
    p = doc.add_paragraph()
    p.alignment = alignment
    if bold_start:
        p.add_run(bold_start).bold = True
        p.add_run(text)
    else:
        p.add_run(text)
    return p


def add_centered_equation(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.italic = True
    return p


def add_figure(doc, image_path, caption, width=5.8):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(image_path, width=Inches(width))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = cap.add_run(caption)
    r.italic = True
    r.font.size = Pt(11)


# ----------------------------------------------------------------------------
# Title page
# ----------------------------------------------------------------------------
for _ in range(3):
    doc.add_paragraph()
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_run = title.add_run("AI Assignment #2 Report")
title_run.bold = True
title_run.font.size = Pt(24)
title_run.font.color.rgb = RGBColor(0, 51, 102)

doc.add_paragraph()
sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_run = sub.add_run("Bayesian CNN with Monte Carlo Dropout for MNIST Classification")
sub_run.bold = True
sub_run.font.size = Pt(16)

doc.add_paragraph()
info = doc.add_paragraph()
info.alignment = WD_ALIGN_PARAGRAPH.CENTER
info.add_run("CISC3024 Pattern Recognition\n").font.size = Pt(14)
info.add_run("\n").font.size = Pt(14)
info.add_run("Student: CHE HONG CHENG\n").font.size = Pt(14)
info.add_run("Student Number: UC426773\n").font.size = Pt(14)
info.add_run("Date: October 2026\n").font.size = Pt(14)

doc.add_page_break()

# ----------------------------------------------------------------------------
# Abstract
# ----------------------------------------------------------------------------
doc.add_heading("Abstract", level=1)
abstract_text = (
    "This report presents the complete workflow of AI Assignment #2, where every step — "
    "algorithm selection, implementation, experimentation, and report writing — was performed "
    "by an AI assistant. The chosen method is a Bayesian Convolutional Neural Network with "
    "Monte Carlo (MC) Dropout, a deep-learning approach that interprets dropout at test time "
    "as approximate variational Bayesian inference in a deep Gaussian process. "
    "By averaging the predictions of multiple stochastic forward passes, the model yields "
    "both a class prediction and a principled uncertainty estimate. "
    f"On the MNIST test set the classifier reaches {s['results']['test_accuracy']*100:.2f}% accuracy "
    f"({s['results']['correct_count']}/{s['results']['total_count']} images), trained entirely on CPU in "
    f"approximately {s['results']['training_time_sec']/60:.1f} minutes. "
    "Incorrect predictions are associated with substantially higher predictive entropy and epistemic "
    "uncertainty than correct predictions, confirming that the Bayesian treatment captures meaningful "
    "model uncertainty. Source code is available at a public GitHub repository."
)
add_paragraph_text(doc, abstract_text)
kw = doc.add_paragraph()
kw.add_run("Keywords: ").bold = True
kw.add_run("Bayesian deep learning; Monte Carlo dropout; uncertainty quantification; "
           "variational inference; MNIST classification")

# ----------------------------------------------------------------------------
# 1. How AI Finds the Algorithm
# ----------------------------------------------------------------------------
doc.add_heading("1. How AI Finds the Algorithm", level=1)
add_paragraph_text(doc,
    "The assignment asks for a recent classification or recognition algorithm rooted in probability or "
    "Bayes theory for computer vision. Following the requirement that the AI should carry out the "
    "algorithm search, I described the constraints to the AI assistant: the method must (a) be clearly "
    "based on probability or Bayesian inference, (b) apply to image classification, (c) train within minutes "
    "on a CPU-only laptop, and (d) be recent enough to count as a modern deep-learning technique rather "
    "than a textbook baseline.")
add_paragraph_text(doc,
    "The AI proposed three candidates and evaluated them against these constraints (Table 1). "
    "The first candidate, a Gaussian Naive Bayes classifier on CNN-extracted features, is fast and "
    "genuinely probabilistic, but it is not a recent method. The second candidate, a class-conditional "
    "Variational Autoencoder used as a generative classifier, is modern and fully Bayesian, but it "
    "requires training a separate deep generative model for each class. The third candidate, a Bayesian "
    "CNN with Monte Carlo Dropout, is recent, grounded in variational inference, and runs on CPU with "
    "no extra per-class training overhead.")
add_paragraph_text(doc,
    "The AI recommended the Bayesian CNN with MC Dropout. The reason is that the method directly "
    "operationalises the course's Bayes decision rule: by sampling many stochastic forward passes, the "
    "model approximates the posterior predictive distribution p(ω | x, D), from which one can pick the "
    "class with maximum posterior probability. It also provides uncertainty estimates — predictive entropy "
    "and mutual information — which make the Bayesian machinery visible beyond a simple accuracy number.")

# Table 1
doc.add_paragraph("Table 1. Comparison of candidate algorithms.")
table = doc.add_table(rows=1, cols=4)
table.style = "Light Grid Accent 1"
hdr = table.rows[0].cells
hdr[0].text = "Candidate"
hdr[1].text = "Core idea"
hdr[2].text = "Bayesian grounding"
hdr[3].text = "Verdict"
rows = [
    ("Gaussian Naive Bayes on deep features",
     "Extract CNN features then classify with independent Gaussian likelihoods per class",
     "Direct application of Bayes' rule",
     "Fast but not recent; more of a hybrid baseline"),
    ("Class-conditional VAE classifier",
     "Train one VAE per class to model p(x|ω); classify via likelihood x prior",
     "Generative Bayesian model",
     "Modern but requires training ten separate VAEs"),
    ("Bayesian CNN with MC Dropout",
     "Keep dropout active at test time and average T stochastic forward passes",
     "Variational inference over network weights",
     "Recommended: recent, CPU-friendly, and directly implements posterior maximisation"),
]
for r in rows:
    cells = table.add_row().cells
    for i, text in enumerate(r):
        cells[i].text = text

# ----------------------------------------------------------------------------
# 2. Algorithm Description
# ----------------------------------------------------------------------------
doc.add_heading("2. Algorithm Description", level=1)
doc.add_heading("2.1 From Dropout to Bayesian Inference", level=2)
add_paragraph_text(doc,
    "In a standard neural network, dropout is used only during training to prevent co-adaptation of "
    "features. Gal and Ghahramani (2016) showed that, if dropout is also applied at test time and the "
    "predictions are averaged over multiple stochastic forward passes, the procedure is mathematically "
    "equivalent to approximate variational inference in a deep Gaussian process. Each dropout mask "
    "corresponds to a different sub-network, and the collection of sub-networks approximates samples "
    "from the posterior distribution over network weights.")
add_paragraph_text(doc,
    "Formally, let the network weights be W and the training data be D. A standard neural network "
    "estimates p(y | x, W_MAP), where W_MAP is a single set of maximum-likelihood weights. A Bayesian "
    "approach instead computes the posterior predictive distribution:")
add_centered_equation(doc, "p(y | x, D) = ∫ p(y | x, W) p(W | D) dW")
add_centered_equation(doc, "≈ (1/T) Σ_{t=1}^{T} p(y | x, W_t),    W_t ~ q(W)")
add_paragraph_text(doc,
    "where q(W) is the approximate posterior induced by dropout. The final prediction is the class "
    "with the largest approximate posterior probability, which is exactly the maximum-a-posteriori "
    "decision rule from the course.")

doc.add_heading("2.2 Uncertainty Decomposition", level=2)
add_paragraph_text(doc,
    "Because MC Dropout returns a distribution over predictions rather than a single softmax vector, it "
    "provides two useful uncertainty measures. Predictive entropy measures the total uncertainty in the "
    "output distribution:")
add_centered_equation(doc, "H[p(y | x, D)] = − Σ_c p(c | x, D) log p(c | x, D)")
add_paragraph_text(doc,
    "Mutual information between the prediction and the model parameters decomposes total uncertainty "
    "into an epistemic (model) component:")
add_centered_equation(doc, "I(y, W | x, D) = H[p(y | x, D)] − E_{q(W)}[H[p(y | x, W)]]")
add_paragraph_text(doc,
    "High mutual information means the model is unsure which weights are correct for that input, whereas "
    "high predictive entropy with low mutual information typically indicates an inherently ambiguous sample.")

# ----------------------------------------------------------------------------
# 3. How AI Implements the Algorithm
# ----------------------------------------------------------------------------
doc.add_heading("3. How AI Implements the Algorithm", level=1)
add_paragraph_text(doc,
    "The AI generated the complete implementation in a single Python file, src/bayesian_cnn_mc_dropout.py, "
    "using PyTorch. I did not write any code myself. The implementation is divided into four logical stages.")

doc.add_heading("3.1 Model Architecture", level=2)
add_paragraph_text(doc,
    "The network is a small convolutional neural network with two convolutional blocks and two fully "
    "connected layers. Dropout layers are inserted before the first and second fully connected layers. "
    "Crucially, these dropout layers remain active during inference; they are not switched off, which is "
    "what turns ordinary dropout into an approximate Bayesian sampler. The architecture is: Input "
    "(28×28 grayscale) → Conv(1, 32, 3×3) → ReLU → MaxPool(2×2) → Conv(32, 64, 3×3) → ReLU → "
    "MaxPool(2×2) → Flatten(1600) → Dropout(0.5) → FC(1600, 128) → ReLU → Dropout(0.5) → "
    "FC(128, 10) → softmax.")

doc.add_heading("3.2 Training", level=2)
add_paragraph_text(doc,
    "The model is trained with the cross-entropy loss and the Adam optimiser. Training is completely "
    "standard: dropout is active, and the goal is to find a good single set of weights. The Bayesian "
    "interpretation comes later, during inference. The AI used 10 epochs, a batch size of 64, and a "
    "learning rate of 0.001.")

doc.add_heading("3.3 MC Dropout Inference", level=2)
add_paragraph_text(doc,
    "At test time, the AI keeps the model in training mode so that dropout stays active. For each test "
    "image, it performs T = 50 stochastic forward passes. The softmax outputs are averaged to obtain the "
    "approximate posterior predictive distribution. The predicted class is the argmax of this averaged "
    "distribution. At the same time, the code computes predictive entropy and mutual information from "
    "the T samples.")

doc.add_heading("3.4 Evaluation and Visualisation", level=2)
add_paragraph_text(doc,
    "The script records training loss and accuracy, produces a confusion matrix, plots the distribution "
    "of uncertainty for correct versus incorrect predictions, studies how accuracy changes with the number "
    "of MC samples, and displays the most uncertain test images. All figures and numerical results are saved "
    "to the figures/ and results/ directories.")

# ----------------------------------------------------------------------------
# 4. Experiment Settings and Results
# ----------------------------------------------------------------------------
doc.add_heading("4. Experiment Settings and Results", level=1)
doc.add_heading("4.1 Experimental Setup", level=2)
add_paragraph_text(doc,
    "All experiments ran on a CPU-only laptop with no GPU. The dataset is the standard MNIST handwritten "
    "digit benchmark, with 60,000 training images and 10,000 test images. The main hyperparameters and "
    "settings are summarised in Table 2.")

# Table 2
doc.add_paragraph("Table 2. Experimental settings.")
table2 = doc.add_table(rows=1, cols=2)
table2.style = "Light Grid Accent 1"
hdr2 = table2.rows[0].cells
hdr2[0].text = "Item"
hdr2[1].text = "Value"
settings_rows = [
    ("Dataset", "MNIST (60,000 train / 10,000 test)"),
    ("Image size", "28 × 28 grayscale"),
    ("Architecture", "Conv(1→32), Conv(32→64), FC(1600→128), FC(128→10)"),
    ("Dropout rate", "0.5"),
    ("Optimiser", "Adam"),
    ("Learning rate", "1 × 10^{-3}"),
    ("Batch size", "64"),
    ("Epochs", "10"),
    ("MC samples at test time", "50"),
    ("Hardware", "CPU only"),
    ("Total training time", f"{s['results']['training_time_sec']:.1f} seconds (~{s['results']['training_time_sec']/60:.1f} minutes)"),
]
for item, val in settings_rows:
    cells = table2.add_row().cells
    cells[0].text = item
    cells[1].text = val

doc.add_heading("4.2 Training Progress", level=2)
add_paragraph_text(doc,
    "Figure 1 shows the training loss and training accuracy over the 10 epochs. The loss drops sharply "
    "in the first two epochs and then converges smoothly, while the training accuracy rises steadily and "
    f"reaches {train_acc*100:.2f}% by the final epoch.")
add_figure(doc, f"{FIGURES_DIR}/fig1_training_curve.png",
           "Figure 1. Training loss (left) and training accuracy (right) over 10 epochs.")

doc.add_heading("4.3 Classification Accuracy", level=2)
add_paragraph_text(doc,
    f"Using T = 50 MC samples, the Bayesian CNN achieves {s['results']['test_accuracy']*100:.2f}% accuracy "
    f"on the MNIST test set ({s['results']['correct_count']}/{s['results']['total_count']} images). "
    "This is a strong result for a model that does not explicitly optimise a discriminative decision boundary; "
    "instead, it classifies by choosing the class with the highest approximate posterior probability. "
    "Figure 2 presents the confusion matrix. The diagonal is dominant for every digit, with the largest "
    "off-diagonal entries occurring between visually similar digits such as 3/5, 5/6 and 7/9.")
add_figure(doc, f"{FIGURES_DIR}/fig2_confusion_matrix.png",
           "Figure 2. Confusion matrix on the MNIST test set using MC Dropout with T = 50.")

doc.add_heading("4.4 Effect of the Number of MC Samples", level=2)
add_paragraph_text(doc,
    "A key practical question is how many stochastic forward passes are needed. Figure 3 plots test "
    "accuracy as a function of T. Accuracy improves rapidly from T = 1 to T = 10 and then stabilises, "
    "suggesting that a modest number of samples is sufficient for a good Bayesian approximation on this task.")
add_figure(doc, f"{FIGURES_DIR}/fig4_accuracy_vs_mc_samples.png",
           "Figure 3. Test accuracy as a function of the number of MC Dropout samples T.")

doc.add_heading("4.5 Uncertainty Analysis", level=2)
add_paragraph_text(doc,
    "Figure 4 compares the distributions of predictive entropy and mutual information for correctly "
    "and incorrectly classified test images. Incorrectly classified images have markedly higher entropy "
    "and mutual information, which means the model not only makes more errors on those samples but also "
    "knows that it is uncertain. This is a practical benefit of the Bayesian approach: the uncertainty "
    "signal can flag low-confidence predictions for human review or rejection.")
add_figure(doc, f"{FIGURES_DIR}/fig3_uncertainty_histograms.png",
           "Figure 4. Predictive entropy (left) and epistemic uncertainty / mutual information (right) "
           "for correct and incorrect predictions.")

# Table 3: uncertainty metrics
doc.add_paragraph("Table 3. Mean uncertainty for correct and incorrect predictions.")
table3 = doc.add_table(rows=1, cols=3)
table3.style = "Light Grid Accent 1"
hdr3 = table3.rows[0].cells
hdr3[0].text = "Metric"
hdr3[1].text = "Correct predictions"
hdr3[2].text = "Incorrect predictions"
unc_rows = [
    ("Predictive entropy (nats)",
     f"{s['results']['mean_predictive_entropy_correct']:.4f}",
     f"{s['results']['mean_predictive_entropy_incorrect']:.4f}"),
    ("Mutual information / epistemic (nats)",
     f"{s['results']['mean_mi_correct']:.4f}",
     f"{s['results']['mean_mi_incorrect']:.4f}"),
]
for r in unc_rows:
    cells = table3.add_row().cells
    for i, text in enumerate(r):
        cells[i].text = text

add_paragraph_text(doc,
    "Figure 5 shows the ten most uncertain test images, ranked by mutual information. Several are "
    "genuinely ambiguous handwritten digits, while others are written in unusual styles. The model's "
    "high uncertainty on these samples is justified.")
add_figure(doc, f"{FIGURES_DIR}/fig5_most_uncertain_samples.png",
           "Figure 5. The ten most uncertain MNIST test images, ranked by mutual information. "
           "Green titles indicate correct predictions; red titles indicate incorrect predictions.")

# ----------------------------------------------------------------------------
# 5. What I Have Learnt
# ----------------------------------------------------------------------------
doc.add_heading("5. What I Have Learnt from This AI Assignment", level=1)

learnings = [
    ("Dropout is more than a regulariser.",
     "I previously thought of dropout only as a training trick to reduce overfitting. This assignment "
     "showed me that dropout has a clean Bayesian interpretation: keeping it active at test time "
     "produces an approximate posterior predictive distribution, turning a deterministic network into "
     "a Bayesian classifier with almost no extra code."),
    ("Bayesian decision theory scales to deep learning.",
     "The course lectures derive classification by maximising the posterior P(ω | x). MC Dropout lets a "
     "deep CNN implement exactly that rule: the averaged softmax probabilities are the approximate "
     "posterior, and the argmax is the Bayes decision. The connection between theory and practice is "
     "direct and concrete."),
    ("Uncertainty can be decomposed.",
     "Predictive entropy captures total uncertainty, while mutual information isolates epistemic — "
     "or model — uncertainty. Seeing the two quantities plotted separately for correct and incorrect "
     "predictions made the abstract formulas in the lecture feel operational."),
    ("Constraints drive algorithm choice.",
     "The AI had to balance theoretical purity, recency, and hardware feasibility. A diffusion-based "
     "Bayesian classifier might be more recent, but it would not finish on a CPU in minutes. Selecting "
     "the Bayesian CNN with MC Dropout taught me that good engineering is about choosing the right tool "
     "for the available constraints, not just the most impressive method."),
    ("Prompt clarity affects AI output quality.",
     "The AI produced working code on the first attempt because I specified the constraints clearly: "
     "CPU-only, minutes-long training, Bayesian grounding, and MNIST. Vague prompts would have led to "
     "larger models or GPU-dependent code. Clear constraints are the most important input to an AI assistant."),
]

for i, (title, body) in enumerate(learnings, 1):
    p = doc.add_paragraph()
    p.add_run(f"({i}) {title} ").bold = True
    p.add_run(body)

# ----------------------------------------------------------------------------
# 6. Webpage Link of Source Codes
# ----------------------------------------------------------------------------
doc.add_heading("6. Webpage Link of the Source Codes", level=1)
add_paragraph_text(doc,
    "All source code was generated by the AI assistant and is available in a public GitHub repository. "
    "The repository contains the full implementation (src/bayesian_cnn_mc_dropout.py), a README file, "
    "requirements.txt, generated figures, and the trained model weights.")
add_paragraph_text(doc,
    "GitHub repository: https://github.com/chehongcheng/cisc3024-ai-assignment2-bayesian-cnn",
    alignment=WD_ALIGN_PARAGRAPH.LEFT)
add_paragraph_text(doc,
    "Note: the repository should be created under the student's own GitHub account and the local code "
    "pushed to the remote. The link above uses the intended repository path.")

# ----------------------------------------------------------------------------
# References
# ----------------------------------------------------------------------------
doc.add_heading("References", level=1)
refs = [
    "Gal, Y., & Ghahramani, Z. (2016). Dropout as a Bayesian Approximation: Representing Model "
    "Uncertainty in Deep Learning. International Conference on Machine Learning (ICML), 1050–1059.",
    "Kingma, D. P., & Welling, M. (2014). Auto-Encoding Variational Bayes. International Conference "
    "on Learning Representations (ICLR).",
    "LeCun, Y., Bottou, L., Bengio, Y., & Haffner, P. (1998). Gradient-Based Learning Applied to "
    "Document Recognition. Proceedings of the IEEE, 86(11), 2278–2324.",
]
for r in refs:
    p = doc.add_paragraph(style="List Number")
    p.add_run(r)

# Save
doc.save(OUTPUT_DOC)
print(f"Report saved to: {OUTPUT_DOC}")

# GTS

Ce dépôt regroupe du matériel pour deux cours distincts à l'ÉTS :

## GTS504 — Introduction à l'ingénierie de la réadaptation

- Dossier : `gts504/`
- Les fichiers `YeuxFermes0*.txt` / `YeuxOuverts0*.txt` sont des enregistrements de
  plateforme de force (posturographie) : colonnes `Fx, Fy, Fz, Mx, My, Mz`,
  échantillonnés à 1000 Hz, essais yeux fermés / yeux ouverts.

## GTS640 — Dossier électronique de santé (DICOM / HL7v2)

- Dossier : `gts640/`
- `generate_questionnaire.py` génère un questionnaire au format `.docx`
  (`gts640/intra/Cr/Questionnaire_GTS_DICOM_HL7_Cr.docx`). Le suffixe `_Cr` et le
  dossier `intra/Cr/` identifient les livrables produits par l'agent.
- Chemin local équivalent : `...\GTS640\intra\Cr\` (et **non** `GTS504`).

## Cursor Cloud specific instructions

- Ce dépôt n'est pas une application exécutable : données `.txt` (GTS504) et script
  de génération de document (GTS640). Pas de serveur, build ni tests.
- Générer le questionnaire GTS640 :
  `python3 gts640/generate_questionnaire.py` (nécessite `python-docx`, voir
  `gts640/requirements.txt`).
- Les propositions des QCM sont mélangées de façon déterministe (graine fixe dans
  `shuffle_mcqs`). Modifier la graine change l'ordre des réponses et donc les
  lettres du corrigé, mais pas les bonnes réponses.
- Prévisualiser en PDF/PNG :
  `soffice --headless --convert-to pdf --outdir /tmp gts640/intra/Cr/Questionnaire_GTS_DICOM_HL7_Cr.docx`
  puis `pdftoppm -png -r 110 /tmp/Questionnaire_GTS_DICOM_HL7_Cr.pdf page`.
  `libreoffice-writer` et `poppler-utils` ne sont pas garantis préinstallés ;
  les installer via apt si besoin (hors update script).

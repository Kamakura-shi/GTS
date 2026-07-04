# GTS

Ce dépôt regroupe des données et du matériel de cours pour le programme GTS.

- Les fichiers `YeuxFermes0*.txt` / `YeuxOuverts0*.txt` sont des enregistrements de
  plateforme de force (posturographie) : colonnes `Fx, Fy, Fz, Mx, My, Mz`,
  échantillonnés à 1000 Hz, essais yeux fermés / yeux ouverts.
- `generate_questionnaire.py` génère un questionnaire au format `.docx`
  (`Questionnaire_GTS_DICOM_HL7.docx`) portant sur les standards DICOM et HL7v2.

## Cursor Cloud specific instructions

- Ce dépôt n'est pas une application exécutable : il contient des données `.txt`
  et un script de génération de document. Il n'y a ni serveur, ni build, ni tests.
- Générer le questionnaire : `python3 generate_questionnaire.py` (nécessite
  `python-docx`, voir `requirements.txt`). Le script contient une banque de
  questions et écrit le `.docx` à la racine du dépôt, corrigé inclus.
- Les propositions des QCM sont mélangées de façon déterministe (graine fixe dans
  `shuffle_mcqs`). Modifier la graine change l'ordre des réponses et donc les
  lettres du corrigé, mais pas les bonnes réponses.
- Prévisualiser en PDF/PNG (utile pour vérifier le rendu) :
  `soffice --headless --convert-to pdf --outdir /tmp Questionnaire_GTS_DICOM_HL7.docx`
  puis `pdftoppm -png -r 110 /tmp/Questionnaire_GTS_DICOM_HL7.pdf page`.
  `libreoffice-writer` et `poppler-utils` ne sont pas garantis préinstallés ;
  les installer via apt si besoin (hors update script).

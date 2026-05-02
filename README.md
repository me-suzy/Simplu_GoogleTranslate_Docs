# Google Translate Docs - Chrome

Proiect separat pentru traducerea documentelor `.doc` / `.docx` din `g:\ARHIVA\C` cu Google Translate Docs.

## Fisiere

- `google_translate_docs_chrome.py` - scriptul principal Chrome
- `run_google_translate_chrome.bat` - pornire normala
- `PowerShell\Start-ChromeDebug.ps1` - porneste Chrome debug pe profilul bun
- `work\` - fisiere intermediare pentru conversie/split
- `downloads\` - fisierele `.doc/.docx/.pdf` traduse si descarcate de la Google Translate; se pastreaza dupa conversia in PDF
- `final_pdf\` - PDF-urile finale `*_FINALIZAT.pdf`
- `logs\` - log detaliat pentru fiecare rulare; fisierul `run_google_translate_chrome_*.log` din folderul principal salveaza si output-ul vazut in CMD
- `state_google_translate_chrome.json` - progres / resume

## Test sigur fara upload

```bat
python google_translate_docs_chrome.py --prepare-only --max-files 3
```

Acest mod doar converteste/spliteaza documentele si arata partile rezultate. Nu porneste Chrome si nu urca nimic pe Google.

## Resume / finalizare fara upload

Daca fisierele traduse au fost deja descarcate, dar PDF-ul final nu a fost creat:

```bat
python google_translate_docs_chrome.py --finalize-existing
```

Acest mod nu porneste Chrome si nu urca nimic. Cauta partile deja descarcate, face merge si creeaza PDF-ul final.

## Rulare reala

```bat
run_google_translate_chrome.bat
```

Rularea normala reia automat progresul din `state_google_translate_chrome.json`: sare peste partile deja descarcate, face PDF pentru documentele complete si continua cu urmatorul `.doc/.docx` din `g:\ARHIVA\C`. In timpul split-ului, scriptul afiseaza fiecare parte terminata si salveaza checkpoint-uri dupa fiecare parte.

Pentru primul test real, cu un singur document:

```bat
run_google_translate_chrome_test_1.bat
```

Scriptul porneste Chrome debug automat daca portul `9222` nu raspunde. Foloseste profilul Chrome:

`C:\Users\necul\AppData\Local\Google\Chrome\User Data\Default`

## Setari utile

- `SIMPLU_GT_MAX_FILES=1` - proceseaza doar primul document, util pentru test
- `SIMPLU_GT_MAX_BYTES=5000000` - limita maxima pentru fiecare upload
- `SIMPLU_GT_MAX_PAGES_PER_PART=400` - limita maxima de pagini pentru fiecare upload
- `SIMPLU_GT_TRANSLATE_WAIT_SEC=60` - asteptare dupa apasarea butonului Translate
- `SIMPLU_GT_BETWEEN_PARTS_SEC=60` - pauza intre partile aceluiasi document
- `SIMPLU_GT_TRANSLATE_ERROR_RETRIES=2` - cate retry-uri face daca Google afiseaza ca fisierul nu poate fi tradus momentan
- `SIMPLU_GT_KEEP_INTERMEDIATE=1` - pastreaza fisierele intermediare dupa succes

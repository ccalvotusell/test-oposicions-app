TEST OPOSICIONS - INSTRUCCIONS RÀPIDES

PER FER SERVIR A L'ORDINADOR
1. Descomprimeix la carpeta.
2. Obre la carpeta test-oposicions-app.
3. Fes doble clic a index.html.
4. Tria mode, tema i nombre de preguntes.
5. Clica Començar.

PREGUNTES AMB ESTRELLA
- Durant el test pots clicar ☆ per marcar una pregunta.
- Queda guardada al navegador d'aquell dispositiu.
- Després pots activar “Només preguntes marcades ⭐”.
- Si esborres les dades del navegador, aquestes marques es poden perdre.

PER ACTUALITZAR PREGUNTES
1. Edita el Word mantenint el format: Tema X_Nom, pregunta numerada, opcions a)-d), resposta correcta en negreta.
2. Executa:
   python3 convert_word_to_json.py test_formated.docx --out preguntes.json --js preguntes.js --report informe_validacio.txt
3. Substitueix preguntes.js dins la carpeta de l'app.

PER COMPARTIR AMB MÒBIL
La manera més fàcil per a gent no tècnica és publicar la carpeta com a web estàtica, per exemple amb Netlify Drop o GitHub Pages, i compartir un enllaç.

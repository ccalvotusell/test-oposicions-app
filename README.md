# Opos Test App

Aplicació web interactiva per a la preparació d'oposicions mitjançant preguntes tipus test organitzades per temes.

---

## Descripció

Aquesta eina ha estat desenvolupada per facilitar l'estudi i l'autoavaluació de temaris d'oposicions. Permet navegar per temes, respondre preguntes tipus test, corregir-les automàticament i identificar preguntes provinents de convocatòries anteriors.

L'aplicació està implementada íntegrament amb HTML, CSS i JavaScript i pot desplegar-se fàcilment en plataformes de hosting estàtic com Netlify o GitHub Pages.

---

## Funcionalitats

- Organització de preguntes per temes.
- Correcció automàtica.
- Preguntes de convocatòries anteriors identificades visualment.
- Suport per explicacions de les respostes.
- Funcionament 100% local o web.
- Conversió automàtica des de documents Word.

---

## Estructura del projecte

```text
.
├── index.html
├── style.css
├── app.js
├── preguntes.js
├── README.md
└── convert_word_to_json.py
```

---

## Actualització de preguntes

Les preguntes es generen a partir d'un document Word amb una estructura determinada.

Exemple:

```text
Tema 1_Nom del tema

1. Text de la pregunta
a) Opció A
b) Opció B
c) Opció C
d) Opció D
```

Després de generar el fitxer `preguntes.js`, només cal substituir-lo dins del projecte.

---

## Desplegament

### GitHub

```bash
git clone <repositori>
```

### Netlify

Connectar el repositori GitHub i desplegar la carpeta arrel del projecte.

---

## Autoria

**Carla Calvó-Tusell**  
https://ccalvotusell.github.io/

© 2026 Carla Calvó-Tusell. Tots els drets reservats.

---

## Condicions d'ús

Aquest programari ha estat desenvolupat per Carla Calvó-Tusell per a ús privat i educatiu.

No es permet:

- La redistribució del codi font.
- La publicació de còpies completes o parcials.
- La modificació i redistribució del projecte.
- L'eliminació dels avisos d'autoria.
- La comercialització del programari o de derivats.

Qualsevol ús, adaptació o redistribució requereix autorització explícita i per escrit de l'autora.

En qualsevol ús autoritzat, s'haurà de mantenir el reconeixement visible de l'autoria original.

---

## Citació

Si aquesta eina s'utilitza com a base per a materials derivats autoritzats, cal incloure:

> Desenvolupat originalment per Carla Calvó-Tusell (2026).

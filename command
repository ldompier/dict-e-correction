# Avec ts-node
npx ts-node colorise_dictee.ts --mapping mapping.csv --compil compilation_mot.csv --out tableau_colore.docx

# (Optionnel) Compiler puis exécuter en Node
npx tsc
node dist/colorise_dictee.js --mapping mapping.csv --compil compilation_mot.csv --out tableau_colore.docx

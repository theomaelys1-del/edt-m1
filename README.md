# EDT M1 MGACL — Groupe 2

Calendrier public de l'année universitaire 2026–2027. Le fichier Excel reste local.

Abonnement : https://raw.githubusercontent.com/theomaelys1-del/edt-m1/main/edt.ics

Apple : webcal://raw.githubusercontent.com/theomaelys1-del/edt-m1/main/edt.ics

```powershell
python -m pip install -r requirements.txt
python generate_ics.py "C:\chemin\emploi-du-temps.xlsx"
```

Le générateur identifie les en-têtes et le Groupe 2, vérifie les jours de semaine,
résout les initiales des enseignants lorsque le tableau permet une correspondance
unique, puis produit `edt.ics` et `edt.state.json`. Le fichier JSON versionné sert
à comparer les mises à jour. Les initiales C. AGUIRRE et F. LE VEN restent telles
quelles car le tableau actuel ne fournit pas de nom complet fiable.

Les UID sont dérivés du début local du cours et du type cours/trajet. Changer la
salle ou l'intervenant conserve l'UID. Déplacer le début supprime l'ancien UID et
crée le nouveau. Les déplacements affichés dans le rapport sont des hypothèses,
sans changement du schéma d'UID. Aucun VTIMEZONE : les heures sont converties en
UTC avec Europe/Paris. Un cours à 9h ou 9h30 reçoit un trajet d'une heure.

Le skill Claude Code `.claude/skills/maj-edt/SKILL.md` décrit la génération, la
comparaison, le commit daté et le push sur main. Les mises à jour sont déclenchées
quand un nouvel Excel est fourni ; aucun fichier source distant n'est surveillé.

S'abonner à l'URL raw de `edt.ics` au lieu d'importer le fichier. Le flux demande
un rafraîchissement horaire (PT1H), mais chaque application décide de son rythme.
Les suppressions et déplacements prennent effet après son rafraîchissement.

Le fichier actuel contient aussi un horaire `09h 12h30` : les deux heures séparées
par un espace sont interprétées comme une plage 09h–12h30.

---
name: maj-edt
description: Mettre à jour le calendrier ICS M1 MGACL Groupe 2 hébergé sur GitHub quand un nouvel Excel d'emploi du temps est fourni.
---

Depuis la racine du dépôt edt-m1, appliquer ce workflow sans redemander confirmation : l'utilisateur autorise la génération, le commit et le push sur main. Demander une clarification seulement si le fichier, le Groupe 2, une date ou une correspondance enseignant est ambigu.

1. Vérifier la branche main et l'état Git, puis récupérer les changements distants avec `git pull --ff-only`. Préserver les modifications étrangères au calendrier ; arrêter si un conflit empêche la mise à jour.
2. Inspecter le nouvel Excel localement avec openpyxl : repérer le bloc Groupe 2 et le tableau des enseignants à droite. Ne jamais ajouter l'Excel au dépôt public. Vérifier que le mapping du générateur correspond à la structure réelle ; adapter le code si nécessaire. Ne jamais inclure le Groupe 1. L'option `--sheet NOM --confirm-group-2` nécessite une vérification préalable que la feuille et le bloc sont bien Groupe 2.
3. Lancer `python generate_ics.py "CHEMIN_EXCEL"`. Garder `edt.state.json`, qui permet la comparaison avec la version précédente. Les dates sont septembre–décembre 2026 puis janvier–août 2027. Une incohérence de jour doit être clarifiée, jamais corrigée silencieusement.
4. Montrer le résumé produit : cours ajoutés, supprimés, déplacements probables (la correspondance entre cours déplacés n'est pas certaine), modifications d'horaire de fin, matière, salle ou intervenant, nombres de cours et trajets, dates extrêmes et cours sans salle. En cas de déplacement avec UID modifié, l'abonnement remplace l'ancien événement par le nouveau au prochain rafraîchissement.
5. Vérifier UTC et CRLF, examiner `git diff`, puis ajouter explicitement `edt.ics`, `edt.state.json` et les éventuelles modifications pertinentes du générateur/skill. Ne pas utiliser `git add .` pour publier des fichiers non examinés. Si aucune modification, ne pas créer de commit vide.
6. Faire un commit daté en Europe/Paris, par exemple `MAJ EDT Groupe 2 — 2026-10-14`, puis `git push origin main`. Vérifier le succès et indiquer le résultat. En cas d'échec de connexion ou de push, préserver les fichiers et expliquer l'étape à reprendre ; ne jamais forcer le push.

Le calendrier public inclut les disciplines, salles et intervenants, mais aucun Excel source. L'adresse raw reste inchangée ; les abonnements actualisent selon leur propre rythme, même si PT1H est demandé.

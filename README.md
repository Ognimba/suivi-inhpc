# Suivi en temps réel – Enquête d'identification des points de vente (INHPC)

Application Streamlit connectée à KoboToolbox (kf.kobotoolbox.org).
Elle relit les données toutes les 15 minutes (réglable) et affiche :

| Onglet | Contenu |
|---|---|
| **Indicateurs** (en haut) | PV identifiés, fiches du jour (écart avec hier), agents actifs, % de fiches avec alerte, taux de réalisation si un objectif est fixé |
| 📈 Avancement | fiches par jour, cumul (avec ligne d'objectif), par département, par type de PV, réalisation par département |
| 👥 Agents | tableau de performance (total, aujourd'hui, hier, jours actifs, moyenne/jour, % alertes, durée médiane, dernière fiche) + tableau agent × jour |
| 🗺️ Carte | position GPS de chaque PV (bleu = sans alerte, rouge = avec alerte) |
| 🛒 Produits | catégories de produits vendues, nombre moyen de produits cochés par type de PV, PV avec très peu de produits |
| ⚠️ Qualité | alertes : nom manquant, GPS manquant / imprécis / hors Congo, téléphone invalide, durée trop courte, type « autre » non précisé, aucune catégorie, doublons possibles |
| 📋 Données | base filtrée + téléchargement Excel / CSV |
| 🔧 Diagnostic | question Kobo retenue pour chaque information (à vérifier une fois) |

Filtres : département, agent, type de PV, période (aujourd'hui, hier, 7 jours, personnalisée).

---

## 1. Récupérer la clé API Kobo
KoboToolbox → icône du compte (en haut à droite) → **Paramètres du compte** → **Sécurité** → **Clé API** → *Afficher* → copier.
L'identifiant du formulaire est déjà réglé : `abVbPCHqtiYSwKP8YAZ4QP` (celui de l'URL de votre projet).

## 2. Tester sur votre ordinateur (facultatif)
```bash
pip install -r requirements.txt
copy .streamlit\secrets.toml.exemple .streamlit\secrets.toml     # Windows
#   puis ouvrir secrets.toml et coller la clé API dans KOBO_TOKEN
streamlit run app.py
```
Le navigateur s'ouvre sur http://localhost:8501.

## 3. Déployer en ligne (gratuit) – Streamlit Community Cloud
1. Créez un compte sur **github.com**, puis un nouveau dépôt (ex. `suivi-inhpc`), de préférence **Private**.
2. Déposez-y les fichiers : `app.py`, `kobo_suivi.py`, `requirements.txt`, `.streamlit/config.toml`
   (bouton *Add file → Upload files*). **Ne déposez jamais `secrets.toml`** (la clé API).
3. Allez sur **share.streamlit.io**, connectez-vous avec GitHub → **Create app** → choisissez le dépôt,
   la branche `main` et le fichier `app.py`.
4. Avant de lancer, ouvrez **Advanced settings → Secrets** et collez le contenu de
   `secrets.toml.exemple` avec votre vraie clé API (et un `MOT_DE_PASSE` si vous voulez protéger l'accès).
5. **Deploy**. Après 2–3 minutes vous obtenez un lien du type `https://suivi-inhpc.streamlit.app` :
   c'est ce lien que vous envoyez au Chef de service (avec le code d'accès si vous en avez mis un).

Pour modifier un réglage plus tard : sur share.streamlit.io → votre app → ⋮ → **Settings → Secrets**.

## 4. Bon à savoir
- **Mise à jour** : la page ouverte se rafraîchit seule toutes les `ACTUALISATION_MINUTES` minutes ;
  le bouton « Actualiser maintenant » force une relecture immédiate de Kobo.
  Une seule lecture Kobo est faite par intervalle, même si plusieurs personnes consultent l'application.
- **Mise en veille** : sur l'offre gratuite, une application sans visite pendant 12 heures se met en veille.
  La première personne qui ouvre le lien voit un bouton « Yes, get this app back up! » : un clic, et elle
  redémarre en moins d'une minute. Les données ne sont pas perdues (elles restent dans Kobo).
- **Sécurité** : la clé API reste dans les secrets Streamlit, jamais dans le code ni dans la page.
  Comme le lien est public par défaut, mettez un `MOT_DE_PASSE`, ou rendez l'app privée
  (Settings → Sharing) et invitez le Chef de service par son adresse e-mail.
- **Nom des agents** : le formulaire n'a pas de question « nom de l'enquêteur » ; l'application utilise alors
  le compte Kobo (`_submitted_by`). Si tous les agents utilisent le même compte, ajoutez une question
  « Nom de l'agent enquêteur » au formulaire (ou indiquez la bonne question dans `[COLONNES]`).
- **Contrôle de la détection** : ouvrez une fois l'onglet 🔧 Diagnostic et vérifiez que chaque
  information pointe vers la bonne question.

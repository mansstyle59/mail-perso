# AffiAds Mail

Application web statique / PWA en français, adaptée à iPhone et hébergeable gratuitement sur GitHub Pages.

## Version actuelle
- Interface responsive et installable comme PWA.
- Identité AffiAds Mail avec icône enveloppe dédiée.
- Liens vers les sites officiels de Proton Mail, Outlook et Gmail.
- Aucun mot de passe, jeton OAuth ou contenu d'e-mail demandé ou enregistré.
- Cache hors ligne de l'interface.
- Sans assistant IA.

## Limite importante
Cette version est un **portail d'accès à la messagerie**, pas encore un client mail intégré : elle ne lit pas et ne synchronise pas les messages et n'envoie pas d'e-mails directement.

GitHub Pages ne fournit pas de serveur de messagerie ni de stockage secret. Un client unifié nécessitera un fournisseur et une intégration API officielle avec OAuth, et probablement un backend sécurisé. Ne place jamais de clé privée ni de secret OAuth dans un dépôt public.

Une adresse personnalisée telle que `toi@tondomaine.fr` nécessite un domaine que tu contrôles et un fournisseur qui gère les e-mails. GitHub Pages ne crée pas cette adresse.

## Déploiement
1. Dans les paramètres du dépôt, ouvre **Settings → Pages**.
2. Sélectionne **GitHub Actions** comme source.
3. Vérifie que le workflow de déploiement est présent dans `.github/workflows/pages.yml`.
4. L'URL habituelle du site sera `https://TON-UTILISATEUR.github.io/mail-perso/`.

## Installation sur iPhone
Ouvre le site publié dans Safari, touche **Partager**, choisis **Sur l'écran d'accueil**, puis **Ajouter**.
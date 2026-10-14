# AffiAds Mail

Application web statique / PWA en français, adaptée à iPhone et hébergeable gratuitement sur GitHub Pages.

## Fonctionnalités préparées
- Connexion Microsoft via MSAL Browser et OAuth Authorization Code + PKCE.
- Lecture de la boîte de réception et affichage des messages.
- Recherche des messages via Microsoft Graph.
- Consultation d'un message et envoi après confirmation.
- Le mot de passe Microsoft n'est jamais demandé ni enregistré par AffiAds Mail.
- Aucun secret client n'est requis dans cette application statique.
- Sans IA.

## Activation Microsoft (nécessaire avant usage)
1. Ouvre https://entra.microsoft.com/ et connecte-toi avec le compte Microsoft créé.
2. Va dans **Identité > Applications > Inscriptions d'applications > Nouvelle inscription**.
3. Donne le nom `AffiAds Mail`.
4. Pour les types de comptes pris en charge, sélectionne **Comptes personnels Microsoft uniquement** (si tu utilises un compte Outlook.com personnel).
5. Dans la plateforme, ajoute une application **Web monopage (SPA)** avec l'URI de redirection exacte du site GitHub Pages : `https://mansstyle59.github.io/mail-perso/`.
6. Enregistre l'application et copie son **ID d’application (client)**. C'est un identifiant public, pas un mot de passe.
7. Dans `app.js`, remplace `REMPLACER_PAR_ID_APPLICATION_MICROSOFT` par cet ID. Ne crée pas de secret client et n'en ajoute aucun au dépôt.
8. Dans **Autorisations API**, ajoute les permissions déléguées Microsoft Graph `User.Read`, `Mail.Read` et `Mail.Send`. Les permissions de messagerie permettent de lire/envoyer les messages du compte connecté.
9. Enregistre les modifications et attends le déploiement GitHub Pages. Ouvre le site, appuie sur **Connecter mon compte Outlook**, puis accepte les autorisations demandées.
10. Teste la lecture, la recherche et l'envoi avec un message de test.

Si l'URL GitHub Pages diffère, utilise l'URL réelle du site à la fois dans la configuration SPA Microsoft et comme URI de redirection. L'URL doit correspondre exactement.

## Sécurité et limites
- L'application utilise MSAL Browser et des jetons d'accès délégués pour le compte connecté. Les jetons sont gérés dans `sessionStorage` par MSAL, et non dans le dépôt.
- N'ajoute jamais de mot de passe, de secret client, de jeton, ni de clé privée à un fichier public ou à GitHub Pages.
- La recherche utilise Microsoft Graph; certains types de recherche peuvent être limités par les capacités de l'API.
- Le contenu HTML des e-mails est affiché dans une iframe isolée avec sandbox. Ne clique pas sur les liens suspects reçus par e-mail.
- Le fournisseur et les limites gratuites de Microsoft peuvent évoluer.
- Cette application statique n'a pas de backend et ne stocke pas les e-mails dans GitHub.

## Déploiement
Dans **Settings → Pages**, sélectionne **GitHub Actions** comme source. Vérifie ensuite le workflow de déploiement dans `.github/workflows/pages.yml`.

## Installation sur iPhone
Ouvre le site publié dans Safari, touche **Partager**, choisis **Sur l'écran d'accueil**, puis **Ajouter**.

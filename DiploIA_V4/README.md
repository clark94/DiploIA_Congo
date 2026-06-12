# DiploIA Congo — v4.0

Plateforme d'aide à la décision diplomatique pour le Cabinet du Ministre des Affaires Étrangères — République du Congo.

## Design

Inspiré du site ERG Predict :
- Fond crème ivoire `#F0EDE6`
- Vert forêt `#1C4A2E`
- Typographie Playfair Display (serif) pour les titres hero
- Sections alternées crème / blanc / vert foncé
- Cards blanches avec ombres légères
- Badges pill avec checkmarks
- Progress bars colorées

## Lancement local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Déploiement Railway (recommandé)

1. Créer un compte sur https://railway.app
2. New Project → Deploy from GitHub repo
3. Start command :
   ```
   streamlit run app.py --server.port=$PORT --server.address=0.0.0.0
   ```
4. Lien public automatique `https://xxx.up.railway.app`

## Déploiement Streamlit Cloud (gratuit)

1. Mettre le dossier sur GitHub (public)
2. https://share.streamlit.io → connecter le repo
3. Déploiement automatique

## Structure

```
DiploIA_V4/
├── app.py
├── requirements.txt
├── README.md
├── assets/
│   └── armoiries.jpg
└── .streamlit/
    └── config.toml
```

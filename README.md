# Générateur PowerPoint MGEN — Streamlit

Application web permettant de transformer un YAML MGEN définitif en PowerPoint, sans installation locale.

## Fonctions

- détection automatique Atelier / Conférence / Webinaire ;
- validation stricte du YAML ;
- choix automatique de la template officielle ;
- génération du `.pptx` ;
- génération d’un rapport technique ;
- traitement temporaire des fichiers.

## Déploiement sur Streamlit Community Cloud

1. Créer un dépôt GitHub.
2. Déposer **tout le contenu de ce dossier à la racine du dépôt**.
3. Ouvrir Streamlit Community Cloud et choisir **Create app**.
4. Sélectionner le dépôt, la branche `main` et le fichier principal `app.py`.
5. Dans les paramètres avancés, conserver Python 3.12.
6. Cliquer sur **Deploy**.

L’application sera disponible à une adresse en `streamlit.app`.

## Lancement local facultatif

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Arborescence importante

- `app.py` : interface web ;
- `mgen_pptx_generator/` : moteur de génération ;
- `resources/templates/` : templates officielles ;
- `resources/schemas/` : schémas YAML ;
- `examples/` : fichiers de démonstration.

## Contrôle avant diffusion

Le moteur effectue les contrôles structurels et techniques. Une vérification visuelle rapide du PowerPoint reste requise avant diffusion.

from flask import Flask, render_template_string

# Importation de vos Blueprints depuis leurs dossiers respectifs
from balance_app.routes import balance_bp
from bourse_app.routes import bourse_bp

app = Flask(__name__)
# Nécéssaire pour les messages "flash" (erreurs de fichiers)
app.secret_key = "cle_secrete_globale_tres_complexe" 

# --- ENREGISTREMENT DES BLUEPRINTS ---
# Tout ce qui est dans balance_bp sera accessible via l'URL /balance
app.register_blueprint(balance_bp, url_prefix='/balance')

# Tout ce qui est dans bourse_bp sera accessible via l'URL /bourse
app.register_blueprint(bourse_bp, url_prefix='/bourse')

# --- PAGE D'ACCUEIL GLOBALE ---
@app.route('/')
def home():
    # Vous pouvez remplacer ce render_template_string par un vrai render_template('home.html') plus tard
    html_content = """
    <!DOCTYPE html>
    <html lang="fr">
    <head>
        <meta charset="UTF-8">
        <title>Portail des Applications</title>
        <style>
            body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f6; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
            .container { background-color: white; padding: 40px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); text-align: center; max-width: 600px; width: 100%; }
            h1 { color: #2c3e50; margin-bottom: 30px; }
            .app-card { background-color: #ecf0f1; padding: 20px; border-radius: 8px; margin-bottom: 20px; text-decoration: none; color: #333; display: block; transition: transform 0.2s, background-color 0.2s; border-left: 5px solid #3498db; }
            .app-card:hover { transform: translateY(-3px); background-color: #e0e6ed; }
            .app-card h2 { margin: 0 0 10px 0; color: #2980b9; }
            .app-card p { margin: 0; color: #7f8c8d; }
            .balance-card { border-left-color: #27ae60; }
            .balance-card h2 { color: #27ae60; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Mes Outils Professionnels</h1>
            
            <a href="/balance" class="app-card balance-card">
                <h2>📊 Générateur de Balance</h2>
                <p>Convertissez vos Livres Journaux en Balances Comptables au format Excel.</p>
            </a>

            <a href="/bourse" class="app-card">
                <h2>📈 Application Bourse</h2>
                <p>Accédez à l'outil d'analyse et de gestion boursière.</p>
            </a>
        </div>
    </body>
    </html>
    """
    return render_template_string(html_content)

if __name__ == '__main__':
    # Démarre le serveur global
    app.run(debug=True, port=5000)
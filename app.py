# -*- coding: utf-8 -*-
"""
app.py
======
Ponto de entrada da aplicação REDEB2B. Inicializa o Flask, carrega variáveis
de ambiente (.env), registra o blueprint da API e as rotas de página, além de
handlers de erro globais.
"""

import os
import sys
import time
import logging
import threading
import webbrowser
from flask import Flask, render_template, jsonify
from dotenv import load_dotenv

load_dotenv(override=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("redeb2b.app")


def _base_dir():
    """Pasta base para localizar templates/static — funciona tanto rodando
    normalmente (python app.py) quanto empacotado com PyInstaller, seja no
    modo --onefile (extrai para uma pasta temporária, sys._MEIPASS) ou
    --onedir (os arquivos ficam ao lado do .exe, também exposto em
    sys._MEIPASS)."""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = _base_dir()


def create_app():
    app = Flask(
        __name__,
        template_folder=os.path.join(BASE_DIR, "templates"),
        static_folder=os.path.join(BASE_DIR, "static"),
    )
    app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-key-troque-em-producao")

    from api import api_bp
    app.register_blueprint(api_bp)

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/analises")
    @app.route("/dashboard")
    def analises():
        return render_template("analises.html")

    @app.errorhandler(404)
    def not_found(_error):
        return jsonify({"ok": False, "error": "Rota não encontrada."}), 404

    @app.errorhandler(500)
    def internal_error(error):
        logger.exception("Erro interno não tratado: %s", error)
        return jsonify({"ok": False, "error": "Erro interno do servidor."}), 500

    return app


app = create_app()


def _abrir_navegador(url, atraso=1.5):
    """Abre o navegador padrão automaticamente após um pequeno atraso (tempo
    do servidor Flask terminar de subir), numa thread separada para não
    travar a inicialização."""
    def _aguardar_e_abrir():
        time.sleep(atraso)
        try:
            webbrowser.open(url)
        except Exception:  # noqa: BLE001
            logger.warning("Não foi possível abrir o navegador automaticamente. Acesse %s manualmente.", url)

    threading.Thread(target=_aguardar_e_abrir, daemon=True).start()


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    debug = os.getenv("FLASK_DEBUG", "true").strip().lower() == "true"
    logger.info("Iniciando REDEB2B Flask app na porta %s (debug=%s)", port, debug)

    # Abre o navegador automaticamente — exceto no processo "filho" que o
    # reloader do Flask cria quando debug=True (senão abriria 2 abas).
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        _abrir_navegador(f"http://localhost:{port}")

    app.run(host="0.0.0.0", port=port, debug=debug)

from flask import Flask
import os

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY','dev')
    from .routes import bp as main_bp
    app.register_blueprint(main_bp)
    return app

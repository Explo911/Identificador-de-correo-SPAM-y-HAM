from flask import *
import kmeans

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("paginaPrincipal.html", result=None)

@app.route("/Prediccion", methods=["POST"])
def results():
    correo = request.form.get("predict")
    resultado = kmeans.choose(correo, "PreprocesamientoSteps")
    return render_template("paginaPrincipal.html", result=resultado)

if __name__ == "__main__":
    app.run()
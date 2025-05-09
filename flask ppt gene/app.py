from flask import Flask, render_template, request, send_file
from cohere_generator import generate_ppt
import os

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    topic = request.form.get("topic")
    model = request.form.get("model")
    num_slides = int(request.form.get("num_slides"))
    theme = request.form.get("theme")  # 'light', 'dark', or 'aesthetic'

    ppt_path = generate_ppt(topic, model, num_slides, theme)

    filename = os.path.basename(ppt_path)

    return send_file(
        ppt_path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

if __name__ == "__main__":
    app.run(debug=True)

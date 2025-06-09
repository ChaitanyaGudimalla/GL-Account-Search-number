from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    if request.method == 'POST':
        number = request.form['number']
        try:
            df = pd.read_excel("data.xlsx")
            match = df[df['Value'] == int(number)]
            if not match.empty:
                result = match.to_html(index=False, classes='result-table')
            else:
                result = "<p class='no-match'>No match found.</p>"
        except Exception as e:
            result = f"<p class='error'>Error: {e}</p>"
    return render_template("index.html", result=result)

if __name__ == '__main__':
    app.run(debug=True, port=5001)

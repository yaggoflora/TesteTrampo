"""
Aplicação web simples para localizar, dentro de uma pasta de rede, subpastas
"Val..." cuja validade esteja vencida há 3 meses ou mais.

Este aplicativo é SOMENTE LEITURA: em nenhum momento arquivos ou pastas são
criados, alterados, movidos ou excluídos.

Como executar:
    pip install -r requirements.txt
    python app.py

Depois acesse: http://127.0.0.1:5000
"""

from flask import Flask, render_template, request, jsonify
from scanner import escanear_pastas

app = Flask(__name__)


@app.route('/')
def pagina_inicial():
    return render_template('index.html')


@app.route('/api/scan')
def api_scan():
    caminho = request.args.get('path', '').strip()

    if not caminho:
        return jsonify({'sucesso': False, 'erro': 'Informe um caminho de pasta.'}), 400

    dados, erro = escanear_pastas(caminho)

    if erro:
        return jsonify({'sucesso': False, 'erro': erro}), 400

    return jsonify({'sucesso': True, **dados})


if __name__ == '__main__':
    # host='127.0.0.1' = acessível apenas nesta máquina (recomendado).
    # Se precisar acessar de outro computador da rede, use host='0.0.0.0',
    # mas com cautela: isso expõe uma ferramenta de leitura de pastas na rede.
    app.run(debug=True, host='127.0.0.1', port=5000)

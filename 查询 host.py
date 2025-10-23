from flask import Flask, request, jsonify
import requests
from flask_cors import CORS  # 仅允许本地前端访问

app = Flask(__name__)
CORS(app, resources={r"/query": {"origins": "http://localhost:*"}})  # 限制仅本地前端可访问

# 360手机号归属地API地址
API_URL = "https://cx.shouji.360.cn/phonearea.php"


@app.route('/query', methods=['GET'])
def query_phone():
    """接收前端请求，转发到360API，返回结果"""
    phone_number = request.args.get('number', '')

    if not phone_number or not phone_number.isdigit() or len(phone_number) != 11:
        return jsonify({"code": -1, "msg": "无效的手机号"}), 400

    try:
        # 转发请求到360API
        response = requests.get(API_URL, params={"number": phone_number}, timeout=5)
        response.raise_for_status()  # 抛出HTTP错误
        return jsonify(response.json())

    except requests.exceptions.RequestException as e:
        return jsonify({"code": -2, "msg": f"API请求失败: {str(e)}"}), 500


if __name__ == '__main__':
    print("本地代理服务启动成功！")
    print("访问地址: http://localhost:5000")
    print("请保持本窗口开启，关闭则服务停止")
    app.run(host='0.0.0.0', port=1029, debug=False)  # 启动服务
from flask import Flask, render_template, jsonify, send_from_directory
import pandas as pd
import os
from pandas import notna
from datetime import datetime

app = Flask(__name__)

# 配置数据文件夹路径
DATA_FOLDER = 'data'
os.makedirs(DATA_FOLDER, exist_ok=True)

def read_data():
    """读取并清洗数据"""
    try:
        # 读取原始数据
        forecast_path = os.path.join(DATA_FOLDER, '预测.xlsx')
        decomposition_path = os.path.join(DATA_FOLDER, '来源分解.xlsx')

        # 检查文件是否存在
        if not os.path.exists(forecast_path):
            print(f"警告: 预测数据文件不存在 - {forecast_path}")
            return pd.DataFrame(), pd.DataFrame()

        if not os.path.exists(decomposition_path):
            print(f"警告: 分解数据文件不存在 - {decomposition_path}")
            return pd.DataFrame(), pd.DataFrame()

        # 读取Excel文件
        forecast_df = pd.read_excel(forecast_path)
        decomposition_df = pd.read_excel(decomposition_path)

        # 列名标准化（根据实际Excel列名调整）
        column_mapping = {
            'date': ['date', '日期', '时间'],
            'forecast': ['预测值', 'forecast', '预测'],
            'actual': ['实际值', 'actual', '实际']
        }

        # 重命名预测数据列
        for target_col, possible_cols in column_mapping.items():
            for col in possible_cols:
                if col in forecast_df.columns:
                    forecast_df.rename(columns={col: target_col}, inplace=True)
                    break

        # 确保必须的列存在
        required_cols = ['date', 'forecast', 'actual']
        for col in required_cols:
            if col not in forecast_df.columns:
                print(f"错误: 预测数据缺少必要的列 - {col}")
                return pd.DataFrame(), pd.DataFrame()

        # 仅过滤无效日期，保留预测值和实际值中的NaN
        forecast_df = forecast_df[notna(forecast_df['date'])]

        # 转换日期为字符串格式
        forecast_df['date'] = forecast_df['date'].astype(str)

        # 确保分解数据包含日期列
        if 'date' not in decomposition_df.columns:
            for col in ['日期', '时间']:
                if col in decomposition_df.columns:
                    decomposition_df.rename(columns={col: 'date'}, inplace=True)
                    break

        if 'date' in decomposition_df.columns:
            decomposition_df['date'] = decomposition_df['date'].astype(str)

        return forecast_df, decomposition_df

    except Exception as e:
        print(f"数据读取失败: {e}")
        return pd.DataFrame(), pd.DataFrame()


@app.route('/')
def index():
    """首页路由"""
    forecast_df, _ = read_data()

    # 获取最新预测值（处理空数据情况）
    latest_forecast = 0.0
    if not forecast_df.empty and not forecast_df['forecast'].empty:
        try:
            latest_forecast = round(forecast_df['forecast'].iloc[-1], 2)
        except IndexError:
            pass

    return render_template(
        'index.html',
        latest_forecast=latest_forecast,
        last_update=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    )


@app.route('/api/gdp_forecast')
def get_gdp_forecast():
    """提供图表数据的API"""
    forecast_df, decomposition_df = read_data()
    response_data = {'line': {}, 'bar': {}}

    # 处理折线图数据
    if not forecast_df.empty:
        # 保留所有预测值，将NaN实际值转为None（前端Chart.js会处理为缺失点）
        response_data['line'] = {
            'labels': forecast_df['date'].tolist(),
            'forecast': forecast_df['forecast'].tolist(),
            'actual': forecast_df['actual'].apply(lambda x: x if not pd.isna(x) else None).tolist()
        }

    # 处理堆叠柱状图数据（保留所有类别，包括全0值）
    if not decomposition_df.empty and 'date' in decomposition_df.columns:
        # 定义允许的类别（与Excel列名一致）
        valid_categories = [
            '生产', '居民消费', '外贸', '货币金融',
            '劳动就业', '交通物流', '财政', '投资', 'GDP', '调查数据'
        ]

        # 筛选出有效的类别列
        available_categories = [col for col in valid_categories if col in decomposition_df.columns]

        # 构建柱状图数据（保留所有类别，即使值全为0）
        bar_data = {
            'labels': decomposition_df['date'].tolist()
        }

        # 添加每个类别的数据
        for category in available_categories:
            bar_data[category] = decomposition_df[category].tolist()

        response_data['bar'] = bar_data

    return jsonify(response_data)


@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory('static', filename)


if __name__ == '__main__':
    app.run()  # 开发环境开启debug模式
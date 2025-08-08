#!/usr/bin/env python3
"""
生成示例Excel表 - 大数据量
包含多个工作表，涵盖不同类型的业务数据
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import string

def generate_employee_data(rows=1000):
    """生成员工数据"""
    departments = ['技术部', '销售部', '市场部', '人事部', '财务部', '运营部', '客服部', '产品部']
    positions = ['经理', '主管', '专员', '助理', '总监', 'VP', 'CEO']
    cities = ['北京', '上海', '广州', '深圳', '杭州', '成都', '武汉', '西安', '南京', '苏州']
    
    data = []
    for i in range(rows):
        employee_id = f"EMP{str(i+1).zfill(6)}"
        name = ''.join(random.choices(string.ascii_uppercase, k=1)) + ''.join(random.choices(string.ascii_lowercase, k=4))
        department = random.choice(departments)
        position = random.choice(positions)
        city = random.choice(cities)
        salary = random.randint(5000, 50000)
        hire_date = datetime.now() - timedelta(days=random.randint(0, 365*5))
        performance_score = round(random.uniform(60, 100), 2)
        
        data.append({
            '员工ID': employee_id,
            '姓名': name,
            '部门': department,
            '职位': position,
            '城市': city,
            '薪资': salary,
            '入职日期': hire_date.strftime('%Y-%m-%d'),
            '绩效评分': performance_score,
            '状态': '在职' if random.random() > 0.1 else '离职'
        })
    
    return pd.DataFrame(data)

def generate_sales_data(rows=2000):
    """生成销售数据"""
    products = ['笔记本电脑', '智能手机', '平板电脑', '智能手表', '无线耳机', '游戏机', '相机', '音响']
    regions = ['华东', '华南', '华北', '华中', '西南', '西北', '东北']
    sales_channels = ['线上商城', '实体店', '代理商', '直销', '电商平台']
    
    data = []
    for i in range(rows):
        order_id = f"ORD{str(i+1).zfill(8)}"
        product = random.choice(products)
        region = random.choice(regions)
        channel = random.choice(sales_channels)
        quantity = random.randint(1, 10)
        unit_price = random.randint(100, 10000)
        total_amount = quantity * unit_price
        order_date = datetime.now() - timedelta(days=random.randint(0, 365*2))
        customer_id = f"CUST{str(random.randint(1, 5000)).zfill(6)}"
        
        data.append({
            '订单ID': order_id,
            '产品名称': product,
            '地区': region,
            '销售渠道': channel,
            '数量': quantity,
            '单价': unit_price,
            '总金额': total_amount,
            '订单日期': order_date.strftime('%Y-%m-%d'),
            '客户ID': customer_id,
            '支付方式': random.choice(['支付宝', '微信', '银行卡', '现金']),
            '订单状态': random.choice(['已完成', '处理中', '已取消', '待付款'])
        })
    
    return pd.DataFrame(data)

def generate_inventory_data(rows=1500):
    """生成库存数据"""
    categories = ['电子产品', '服装鞋帽', '家居用品', '食品饮料', '图书音像', '运动户外', '美妆护肤', '母婴用品']
    suppliers = ['供应商A', '供应商B', '供应商C', '供应商D', '供应商E', '供应商F', '供应商G', '供应商H']
    warehouses = ['北京仓', '上海仓', '广州仓', '深圳仓', '杭州仓', '成都仓']
    
    data = []
    for i in range(rows):
        sku = f"SKU{str(i+1).zfill(8)}"
        category = random.choice(categories)
        supplier = random.choice(suppliers)
        warehouse = random.choice(warehouses)
        current_stock = random.randint(0, 1000)
        min_stock = random.randint(10, 100)
        max_stock = random.randint(200, 2000)
        unit_cost = round(random.uniform(10, 1000), 2)
        last_updated = datetime.now() - timedelta(days=random.randint(0, 30))
        
        data.append({
            'SKU编码': sku,
            '产品名称': f"{category}产品{i+1}",
            '类别': category,
            '供应商': supplier,
            '仓库': warehouse,
            '当前库存': current_stock,
            '最低库存': min_stock,
            '最高库存': max_stock,
            '单位成本': unit_cost,
            '库存价值': current_stock * unit_cost,
            '最后更新': last_updated.strftime('%Y-%m-%d'),
            '库存状态': '充足' if current_stock > min_stock else '不足'
        })
    
    return pd.DataFrame(data)

def generate_financial_data(rows=1200):
    """生成财务数据"""
    account_types = ['现金', '银行存款', '应收账款', '存货', '固定资产', '应付账款', '预收账款', '长期借款']
    departments = ['技术部', '销售部', '市场部', '人事部', '财务部', '运营部']
    
    data = []
    for i in range(rows):
        transaction_id = f"TXN{str(i+1).zfill(8)}"
        account_type = random.choice(account_types)
        department = random.choice(departments)
        amount = round(random.uniform(100, 100000), 2)
        transaction_type = random.choice(['收入', '支出'])
        transaction_date = datetime.now() - timedelta(days=random.randint(0, 365))
        
        data.append({
            '交易ID': transaction_id,
            '账户类型': account_type,
            '部门': department,
            '金额': amount,
            '交易类型': transaction_type,
            '交易日期': transaction_date.strftime('%Y-%m-%d'),
            '描述': f"{transaction_type} - {account_type}",
            '审核状态': random.choice(['已审核', '待审核', '已拒绝']),
            '经办人': f"用户{random.randint(1, 100)}",
            '备注': f"备注信息{i+1}"
        })
    
    return pd.DataFrame(data)

def generate_customer_data(rows=800):
    """生成客户数据"""
    customer_types = ['个人客户', '企业客户', 'VIP客户', '普通客户']
    industries = ['科技', '金融', '教育', '医疗', '制造', '零售', '服务', '其他']
    
    data = []
    for i in range(rows):
        customer_id = f"CUST{str(i+1).zfill(6)}"
        customer_type = random.choice(customer_types)
        industry = random.choice(industries)
        total_orders = random.randint(1, 100)
        total_spent = round(random.uniform(1000, 100000), 2)
        registration_date = datetime.now() - timedelta(days=random.randint(0, 365*3))
        last_order_date = datetime.now() - timedelta(days=random.randint(0, 365))
        
        data.append({
            '客户ID': customer_id,
            '客户名称': f"客户{i+1}",
            '客户类型': customer_type,
            '行业': industry,
            '联系电话': f"1{random.randint(3000000000, 3999999999)}",
            '邮箱': f"customer{i+1}@example.com",
            '总订单数': total_orders,
            '总消费金额': total_spent,
            '注册日期': registration_date.strftime('%Y-%m-%d'),
            '最后订单日期': last_order_date.strftime('%Y-%m-%d'),
            '客户等级': 'A' if total_spent > 50000 else 'B' if total_spent > 10000 else 'C',
            '状态': '活跃' if random.random() > 0.2 else '非活跃'
        })
    
    return pd.DataFrame(data)

def generate_project_data(rows=600):
    """生成项目数据"""
    project_types = ['产品开发', '系统集成', '咨询服务', '培训项目', '维护服务', '定制开发']
    statuses = ['规划中', '进行中', '已完成', '已暂停', '已取消']
    priorities = ['高', '中', '低']
    
    data = []
    for i in range(rows):
        project_id = f"PRJ{str(i+1).zfill(6)}"
        project_type = random.choice(project_types)
        status = random.choice(statuses)
        priority = random.choice(priorities)
        budget = round(random.uniform(10000, 1000000), 2)
        start_date = datetime.now() - timedelta(days=random.randint(0, 365*2))
        end_date = start_date + timedelta(days=random.randint(30, 365))
        progress = random.randint(0, 100)
        
        data.append({
            '项目ID': project_id,
            '项目名称': f"项目{i+1}",
            '项目类型': project_type,
            '状态': status,
            '优先级': priority,
            '预算': budget,
            '开始日期': start_date.strftime('%Y-%m-%d'),
            '结束日期': end_date.strftime('%Y-%m-%d'),
            '进度': f"{progress}%",
            '项目经理': f"经理{random.randint(1, 50)}",
            '团队成员数': random.randint(3, 20),
            '实际花费': round(budget * random.uniform(0.5, 1.2), 2),
            '风险等级': random.choice(['低', '中', '高'])
        })
    
    return pd.DataFrame(data)

def main():
    """主函数 - 生成Excel文件"""
    print("开始生成示例Excel数据...")
    
    # 创建Excel写入器
    with pd.ExcelWriter('sample_large_excel_data.xlsx', engine='openpyxl') as writer:
        
        # 生成员工数据
        print("生成员工数据...")
        employee_df = generate_employee_data(1000)
        employee_df.to_excel(writer, sheet_name='员工信息', index=False)
        
        # 生成销售数据
        print("生成销售数据...")
        sales_df = generate_sales_data(2000)
        sales_df.to_excel(writer, sheet_name='销售数据', index=False)
        
        # 生成库存数据
        print("生成库存数据...")
        inventory_df = generate_inventory_data(1500)
        inventory_df.to_excel(writer, sheet_name='库存管理', index=False)
        
        # 生成财务数据
        print("生成财务数据...")
        financial_df = generate_financial_data(1200)
        financial_df.to_excel(writer, sheet_name='财务记录', index=False)
        
        # 生成客户数据
        print("生成客户数据...")
        customer_df = generate_customer_data(800)
        customer_df.to_excel(writer, sheet_name='客户信息', index=False)
        
        # 生成项目数据
        print("生成项目数据...")
        project_df = generate_project_data(600)
        project_df.to_excel(writer, sheet_name='项目管理', index=False)
        
        # 生成汇总数据
        print("生成汇总数据...")
        summary_data = {
            '数据类别': ['员工信息', '销售数据', '库存管理', '财务记录', '客户信息', '项目管理'],
            '记录数量': [len(employee_df), len(sales_df), len(inventory_df), len(financial_df), len(customer_df), len(project_df)],
            '生成时间': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')] * 6,
            '数据状态': ['正常'] * 6
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='数据汇总', index=False)
    
    print(f"Excel文件生成完成！")
    print(f"文件名: sample_large_excel_data.xlsx")
    print(f"总记录数: {len(employee_df) + len(sales_df) + len(inventory_df) + len(financial_df) + len(customer_df) + len(project_df)}")
    print(f"工作表数: 7个")
    print(f"数据表: 员工信息({len(employee_df)}行), 销售数据({len(sales_df)}行), 库存管理({len(inventory_df)}行), 财务记录({len(financial_df)}行), 客户信息({len(customer_df)}行), 项目管理({len(project_df)}行)")

if __name__ == "__main__":
    main() 
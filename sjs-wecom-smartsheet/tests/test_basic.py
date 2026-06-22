#!/usr/bin/env python3
"""
基础测试脚本
"""

import os
import sys
import json

# 添加 scripts 目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))


def test_data_file_exists():
    """测试数据文件是否存在"""
    data_file = os.path.expanduser(
        "~/work_space/ai/skill/sjs-skills/sjs-wecom-smartsheet/output/wecom_project_data.json"
    )
    if os.path.exists(data_file):
        print("✅ 数据文件存在")
        return True
    else:
        print("⚠️  数据文件不存在（需要先运行读取脚本）")
        return False


def test_data_structure():
    """测试数据结构"""
    data_file = os.path.expanduser(
        "~/work_space/ai/skill/sjs-skills/sjs-wecom-smartsheet/output/wecom_project_data.json"
    )
    if not os.path.exists(data_file):
        print("⚠️  数据文件不存在，跳过结构测试")
        return False

    with open(data_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 检查必要字段
    required_fields = ['projects', 'stats', 'timestamp']
    for field in required_fields:
        if field not in data:
            print(f"❌ 缺少必要字段: {field}")
            return False

    print("✅ 数据结构正确")
    return True


def test_analyzer():
    """测试分析器"""
    try:
        from wecom数据分析 import WeComAnalyzer
        print("✅ 分析器模块导入成功")
        return True
    except ImportError as e:
        print(f"❌ 分析器模块导入失败: {e}")
        return False


def main():
    """运行所有测试"""
    print("=== 运行基础测试 ===\n")

    tests = [
        test_data_file_exists,
        test_data_structure,
        test_analyzer
    ]

    results = []
    for test in tests:
        try:
            result = test()
            results.append(result)
        except Exception as e:
            print(f"❌ 测试失败: {e}")
            results.append(False)
        print()

    # 打印测试结果
    passed = sum(results)
    total = len(results)
    print(f"=== 测试完成: {passed}/{total} 通过 ===")

    return all(results)


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)

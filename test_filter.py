r'''
Author: jianquan.liu 
Date: 2025-07-02 14:15:41
LastEditors: jianquan.liu 
LastEditTime: 2025-07-03 11:45:04
FilePath: \McuTest\test_filter.py
Description: 

Copyright (c) 2025 by carizon, All Rights Reserved. 
'''

import json
import os
import fnmatch
import logging
from typing import List, Dict, Any, Optional
import pytest

class TestCaseFilter:
    """测试用例过滤器类"""
    
    def __init__(self, config_file: str = "test_config.json"):
        """
        初始化测试过滤器
        
        Args:
            config_file: JSON配置文件路径
        """
        self.config_file = config_file
        self.config = self._load_config()
        
    def _load_config(self) -> Dict[str, Any]:
        """加载JSON配置文件"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                logging.warning(f"配置文件 {self.config_file} 不存在，使用默认配置")
                return self._get_default_config()
        except Exception as e:
            logging.error(f"加载配置文件失败: {e}")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """获取默认配置"""
        return {
            "test_execution_config": {
                "description": "主配置，控制全局测试执行行为。include模式： exclude 模式",
                "execution_mode": "include",
                "test_cases": {
                    "description": "精确指定要执行的测试用例列表，只进行 exclude case 过滤(高优先级，使能之后不再通过 test_filters 过滤case，直接使用精确指定的 case),和 excluded_cases 可用搭配使用",
                    "enabled": False,
                    "cases": []
                },
                "test_filters": {
                    "description": "包含性过滤器，仅执行满足匹配条件的测试用例。和 excluded_tests 搭配使用",
                    "by_file": {
                        "description": "按文件名过滤，仅执行指定文件中的测试。",
                        "enabled": False,
                        "files": []
                    },
                    "by_class": {
                        "description": "按类名过滤，仅执行指定类的测试。",
                        "enabled": False,
                        "classes": []
                    },
                    "by_function": {
                        "description": "按函数名过滤，仅执行指定函数的测试。",
                        "enabled": False,
                        "functions": []
                    },
                    "by_pattern": {
                        "description": "按re包含过滤，仅执行匹配re的测试。",
                        "enabled": False,
                        "patterns": []
                    }
                },
                "excluded_tests": {
                    "description": "模糊排除，按文件、类、函数或模式批量排除测试。",
                    "enabled": False,
                    "by_file": {
                        "description": "按文件名排除测试。",
                        "enabled": False,
                        "files": []
                    },
                    "by_class": {
                        "description": "按类名排除测试。",
                        "enabled": False,
                        "classes": []
                    },
                    "by_function": {
                        "description": "按函数名排除测试。",
                        "enabled": False,
                        "functions": []
                    },
                    "by_pattern": {
                        "description": "按re模式排除测试。",
                        "enabled": False,
                        "patterns": []
                    }
                },
                "excluded_cases": {
                    "description": "精确排除指定的测试用例,只对精确指定的执行case有效。",
                    "enabled": False,
                    "cases": []
                }
            }
        }
    
    def generate_pytest_args(self, base_args: List[str], config_name: str = "test_execution_config") -> List[str]:
        """
        根据配置生成pytest参数
        
        Args:
            base_args: 基础pytest参数
            config_name: 配置名称
            
        Returns:
            完整的pytest参数列表
        """
        try:
            if config_name in self.config.get("test_suite_configs", {}):
                config = self.config["test_suite_configs"][config_name]
            else:
                config = self.config.get(config_name, {})
            
            execution_mode = config.get("execution_mode", "include")
            
            # 优先处理精确的测试用例列表,只进行exclude case 过滤
            if config.get("test_cases", {}).get("enabled", False):
                print(f"使用精确测试用例列表: {config_name}")
                return self._generate_args_from_case_list(base_args, config, execution_mode)
            
            # 如果没有精确的用例列表，使用的过滤方式
            test_filters = config.get("test_filters", {})
            excluded_tests = config.get("excluded_tests", {})
            excluded_cases = config.get("excluded_cases", {})
            
            # 构建测试选择表达式
            expressions = []
            
            # 处理包含模式
            if execution_mode == "include":
                include_expressions = self._build_include_expressions(test_filters)
                if include_expressions:
                    expressions.extend(include_expressions)
            
            # 处理排除模式
            exclude_expressions = self._build_exclude_expressions(excluded_tests)
            if excluded_cases:
                case_exclude_expressions = self._build_case_exclude_expressions(excluded_cases)
                exclude_expressions.extend(case_exclude_expressions)
            if exclude_expressions:
                expressions.extend(exclude_expressions)
            
            # 构建最终参数
            final_args = base_args.copy()
            
            if expressions:
                # 使用 -k 参数进行过滤
                filter_expr = " and ".join(f"({expr})" for expr in expressions)
                final_args.extend(["-k", filter_expr])
                print(f"应用测试过滤器: {filter_expr}")
            
            return final_args
            
        except Exception as e:
            logging.error(f"生成pytest参数失败: {e}")
            return base_args
    
    def _generate_args_from_case_list(self, base_args: List[str], config: Dict[str, Any], execution_mode: str) -> List[str]:
        """
        从精确的测试用例列表生成pytest参数
        
        Args:
            base_args: 基础参数
            config: 配置信息
            execution_mode: 执行模式
            
        Returns:
            包含精确测试用例的参数列表
        """
        # 移除基础参数中的测试目录，因为要精确指定测试用例
        filtered_args = []
        deleted_dirs = []
        for arg in base_args:
            if not (arg in ("Testcases", "Test_yace", "secure_pressure_cases") \
                or arg.endswith("/") or arg.endswith("\\") or \
                (not arg.startswith("-") and "/" in arg) or (not arg.startswith("-") and "\\" in arg)):
                filtered_args.append(arg)
            else:
                deleted_dirs.append(arg)
                logging.debug(f"精确指定测试用例,移除测试目录参数: {arg}")
        
        if execution_mode == "include":
            # 包含模式：只执行指定的测试用例
            test_cases = config.get("test_cases", {}).get("cases", [])
            
            # 应用排除用例
            excluded_cases_config = config.get("excluded_cases", {})
            if excluded_cases_config.get("enabled", False):
                excluded_cases = set(excluded_cases_config.get("cases", []))
                test_cases = [case for case in test_cases if case not in excluded_cases]

            if test_cases:
                filtered_args.extend(test_cases)
                print(f"精确执行 {len(test_cases)} 个测试用例")
                for case in test_cases:
                    print(f"  - {case}")
            else:
                logging.warning("没有指定要执行的测试用例")
                
        else:
            # 排除模式：执行除指定用例外的所有测试
            excluded_cases_config = config.get("excluded_cases", {})
            if excluded_cases_config.get("enabled", False):
                excluded_cases = excluded_cases_config.get("cases", [])
                
                # 构建排除表达式
                if excluded_cases:
                    exclude_patterns = []
                    for case in excluded_cases:
                        # 从完整的nodeid中提取关键部分进行匹配
                        parts = case.split("::")
                        if len(parts) >= 3:
                            # 格式: file::class::method
                            file_part = parts[0].split("/")[-1].replace(".py", "")
                            class_part = parts[1]
                            method_part = parts[2].split("[")[0]  # 移除参数化部分
                            exclude_patterns.append(f"not ({file_part} and {class_part} and {method_part})")
                        elif len(parts) >= 2:
                            # 格式: file::method
                            file_part = parts[0].split("/")[-1].replace(".py", "")
                            method_part = parts[1].split("[")[0]
                            exclude_patterns.append(f"not ({file_part} and {method_part})")
                            
                    '''
                    -k 参数用于按照测试名称、类名或文件名的模式来选择要运行的测试。支持 Python 表达式语法，可以进行灵活的测试筛选。
                    # 运行包含 "login" 或 "logout" 的测试
                    pytest -k "login or logout"
                    # 运行包含 "user" 且包含 "create" 的测试
                    pytest -k "user and create"
                    # 运行包含 "test" 但不包含 "slow" 的测试
                    pytest -k "test and not slow"
                    '''
                    if exclude_patterns:
                        filter_expr = " and ".join(exclude_patterns)
                        filtered_args.extend(["-k", filter_expr])
                        print(f"排除 {len(excluded_cases)} 个测试用例")
                        
            # 在排除模式下，需要添加回测试目录
            for arg in deleted_dirs:
                if arg not in filtered_args:
                    filtered_args.append(arg)
                
        return filtered_args
    
    def _build_include_expressions(self, test_filters: Dict[str, Any]) -> List[str]:
        """构建包含表达式"""
        expressions = []
    
        # 按文件过滤
        if test_filters.get("by_file", {}).get("enabled", False):
            files = [f for f in test_filters["by_file"].get("files", []) if f and f.strip()]
            if files:
                # 使用 pytest 支持的表达式语法
                patterns = [file.replace('.py', '') for file in files]
                if patterns:
                    expressions.append("(" + " or ".join(patterns) + ")")
    
        # 按类过滤
        if test_filters.get("by_class", {}).get("enabled", False):
            classes = [c for c in test_filters["by_class"].get("classes", []) if c and c.strip()]
            if classes:
                # 处理可能包含方法的类名格式
                patterns = []
                for class_name in classes:
                    if "::" in class_name:
                        class_name = class_name.split("::")[0]
                    if class_name.strip():
                        patterns.append(class_name)
                if patterns:
                    expressions.append("(" + " or ".join(patterns) + ")")
    
        # 按函数过滤
        if test_filters.get("by_function", {}).get("enabled", False):
            functions = [f for f in test_filters["by_function"].get("functions", []) if f and f.strip()]
            if functions:
                patterns = []
                for func in functions:
                    # 处理参数化测试用例
                    base_func = func.split("[")[0] if "[" in func else func
                    if base_func.strip():
                        patterns.append(base_func)
                if patterns:
                    expressions.append("(" + " or ".join(patterns) + ")")
    
        # 按模式过滤
        if test_filters.get("by_pattern", {}).get("enabled", False):
            patterns = [p for p in test_filters["by_pattern"].get("patterns", []) if p and p.strip()]
            if patterns:
                # 将通配符模式转换为 pytest 表达式
                pattern_exprs = []
                for pattern in patterns:
                    # 移除*号，转换为简单的包含匹配
                    parts = [p for p in pattern.split("*") if p and p.strip()]
                    if parts:
                        pattern_exprs.extend(parts)
                if pattern_exprs:
                    expressions.append("(" + " or ".join(pattern_exprs) + ")")
    
        return expressions
    
    def _build_exclude_expressions(self, excluded_tests: Dict[str, Any]) -> List[str]:
        """构建排除表达式"""
        expressions = []

        # 按文件排除
        if excluded_tests.get("by_file", False).get("enabled", False):  
            excluded_files = excluded_tests.get("by_file", [])
            if excluded_files:
                patterns = [f"not {file.replace('.py', '')}" for file in excluded_files]
                if patterns:
                    expressions.append(" and ".join(patterns))

        # 按类排除
        if excluded_tests.get("by_class", False).get("enabled", False):
            excluded_classes = [c for c in excluded_tests.get("by_class", []) if c.strip()]
            if excluded_classes:
                patterns = [f"not {class_name}" for class_name in excluded_classes]
                if patterns:
                    expressions.append(" and ".join(patterns))

        # 按函数排除
        if excluded_tests.get("by_function", False).get("enabled", False):
            excluded_functions = excluded_tests.get("by_function", [])
            if excluded_functions:
                patterns = []
                for func in excluded_functions:
                    base_func = func.split("[")[0] if "[" in func else func
                    patterns.append(f"not {base_func}")
                if patterns:
                    expressions.append(" and ".join(patterns))

        # 按re模式排除
        if excluded_tests.get("by_pattern", False).get("enabled", False):
            excluded_patterns = excluded_tests.get("by_pattern", [])
            if excluded_patterns:
                pattern_exprs = []
                for pattern in excluded_patterns:
                    parts = [p for p in pattern.split("*") if p]
                    if parts:
                        pattern_exprs.extend([f"not {part}" for part in parts])
                if pattern_exprs:
                    expressions.append(" and ".join(pattern_exprs))

        return expressions
    
    def _build_case_exclude_expressions(self, excluded_cases: Dict[str, Any]) -> List[str]:
        """构建用例级别的排除表达式"""
        expressions = []
        
        if excluded_cases.get("enabled", False):
            cases = [c for c in excluded_cases.get("cases", []) if c and c.strip()]
            if cases:
                case_patterns = []
                for case in cases:
                    # 从完整的nodeid中提取关键信息用于匹配
                    parts = [p for p in case.split("::") if p and p.strip()]
                    if len(parts) >= 3:
                        # 格式: file::class::method[param]
                        file_part = parts[0].split("/")[-1].replace(".py", "").strip()
                        class_part = parts[1].strip()
                        method_part = parts[2].split("[")[0].strip()  # 移除参数化部分
                        param_part = ""
                        
                        if "[" in parts[2] and "]" in parts[2]:
                            param_part = parts[2].split("[")[1].split("]")[0].strip()
                        
                        pattern_parts = []
                        if file_part:
                            pattern_parts.append(f"not {file_part}")
                        if class_part:
                            pattern_parts.append(f"not {class_part}")
                        if method_part:
                            pattern_parts.append(f"not {method_part}")
                        if param_part:
                            pattern_parts.append(f"not [{param_part}]")
                            
                        if pattern_parts:
                            case_patterns.append("(" + " and ".join(pattern_parts) + ")")
                            
                    elif len(parts) >= 2:
                        # 格式: file::method
                        file_part = parts[0].split("/")[-1].replace(".py", "").strip()
                        method_part = parts[1].split("[")[0].strip()
                        
                        pattern_parts = []
                        if file_part:
                            pattern_parts.append(f"not {file_part}")
                        if method_part:
                            pattern_parts.append(f"not {method_part}")
                            
                        if pattern_parts:
                            case_patterns.append("(" + " and ".join(pattern_parts) + ")")
                
                if case_patterns:
                    expressions.append(" and ".join(case_patterns))
        
        return expressions
    
    def get_available_configs(self) -> List[str]:
        """获取可用的配置名称"""
        configs = ["test_execution_config"]
        suite_configs = self.config.get("test_suite_configs", {})
        configs.extend(suite_configs.keys())
        return configs
    
    def print_config_info(self, config_name: str = "test_execution_config"):
        """打印配置信息"""
        if config_name in self.config.get("test_suite_configs", {}):
            config = self.config["test_suite_configs"][config_name]
        else:
            config = self.config.get(config_name, {})
        
        print(f"\n=== 测试配置: {config_name} ===")
        print(f"描述: {config.get('description', 'N/A')}")
        print(f"执行模式: {config.get('execution_mode', 'N/A')}")
        
        # 显示精确的测试用例列表
        test_cases_config = config.get("test_cases", {})
        if test_cases_config.get("enabled", False):
            cases = test_cases_config.get("cases", [])
            print(f"\n精确指定的测试用例 ({len(cases)} 个):")
            for i, case in enumerate(cases, 1):
                print(f"  {i:2d}. {case}")
        
        # 显示过滤器
        test_filters = config.get("test_filters", {})
        enabled_filters = {k: v for k, v in test_filters.items() if isinstance(k, dict) and k.get("enabled", False)}
        if enabled_filters:
            print("\n启用的过滤器:")
            for filter_type, filter_config in enabled_filters.items():
                print(f"  - {filter_type}: {filter_config}")
        
        # 显示排除的测试用例
        excluded_cases_config = config.get("excluded_cases", {})
        if excluded_cases_config.get("enabled", False):
            excluded_cases = excluded_cases_config.get("cases", [])
            print(f"\n排除的测试用例 ({len(excluded_cases)} 个):")
            for i, case in enumerate(excluded_cases, 1):
                print(f"  {i:2d}. {case}")
        
        # 显示排除
        excluded_tests = config.get("excluded_tests", {})
        if excluded_tests:
            print(f"\n排除规则: {excluded_tests}")
    
    def collect_available_test_cases(self, test_dir: str) -> List[str]:
        """
        收集可用的测试用例列表
        
        Args:
            test_dir: 测试目录
            
        Returns:
            所有可用测试用例的nodeid列表
        """
        collected = []

        class CollectorPlugin:
            def pytest_collection_finish(self, session):
                for item in session.items:
                    collected.append(item.nodeid)
                    
        pytest.main([test_dir, "--collect-only", "-q"], plugins=[CollectorPlugin()])
        
        return collected
    
    def validate_test_cases(self, config_name: str, test_dir: str) -> Dict[str, Any]:
        """
        验证配置中的测试用例是否存在
        
        Args:
            config_name: 配置名称
            test_dir: 测试目录
            
        Returns:
            验证结果
        """
        if config_name in self.config.get("test_suite_configs", {}):
            config = self.config["test_suite_configs"][config_name]
        else:
            config = self.config.get(config_name, {})
        
        # 收集所有可用的测试用例
        available_cases = self.collect_available_test_cases(test_dir)
        available_set = set(available_cases)
        
        result = {
            "config_name": config_name,
            "total_available": len(available_cases),
            "valid_cases": [],
            "invalid_cases": [],
            "valid_excluded": [],
            "invalid_excluded": []
        }
        
        # 验证包含的测试用例
        test_cases_config = config.get("test_cases", {})
        if test_cases_config.get("enabled", False):
            specified_cases = test_cases_config.get("cases", [])
            for case in specified_cases:
                if case in available_set:
                    result["valid_cases"].append(case)
                else:
                    result["invalid_cases"].append(case)
        
        # 验证排除的测试用例
        excluded_cases_config = config.get("excluded_cases", {})
        if excluded_cases_config.get("enabled", False):
            excluded_cases = excluded_cases_config.get("cases", [])
            for case in excluded_cases:
                if case in available_set:
                    result["valid_excluded"].append(case)
                else:
                    result["invalid_excluded"].append(case)
        
        return result
    
    def export_available_cases_to_json(self, test_dir: str, output_file: str = "available_test_cases.json"):
        """
        导出所有可用的测试用例到JSON文件，便于配置时参考
        
        Args:
            test_dir: 测试目录
            output_file: 输出文件名
        """
        collected_details = []

        class DetailedCollectorPlugin:
            def pytest_collection_finish(self, session):
                for item in session.items:
                    test_info = {
                        "nodeid": item.nodeid,
                        "name": item.name,
                        "file": item.location[0] if item.location else "",
                        "line": item.location[1] if item.location else 0,
                        "class": getattr(item.parent, 'name', '') if hasattr(item, 'parent') else '',
                        "function": item.name,
                        "markers": [marker.name for marker in item.iter_markers()],
                        "description": getattr(item.function, '__doc__', '') if hasattr(item, 'function') else ''
                    }
                    collected_details.append(test_info)

        pytest.main([test_dir, "--collect-only", "-q"], plugins=[DetailedCollectorPlugin()])
        
        # 按文件分组
        by_file = {}
        for test in collected_details:
            file_path = test["file"]
            if file_path not in by_file:
                by_file[file_path] = []
            by_file[file_path].append(test)
        
        export_data = {
            "collection_info": {
                "test_directory": test_dir,
                "total_cases": len(collected_details),
                "total_files": len(by_file),
                "generated_at": __import__('datetime').datetime.now().isoformat()
            },
            "by_file": by_file,
            "all_cases": [test["nodeid"] for test in collected_details],
            "sample_config": {
                "description": "示例配置 - 复制需要的测试用例到你的配置中",
                "execution_mode": "include",
                "test_cases": {
                    "enabled": True,
                    "cases": [test["nodeid"] for test in collected_details[:5]]  # 前5个作为示例
                }
            }
        }
        
        with open(output_file, 'w', encoding='utf-8') as f:
            import json
            json.dump(export_data, f, indent=2, ensure_ascii=False)
        
        print(f"可用测试用例已导出到: {output_file}")
        print(f"共发现 {len(collected_details)} 个测试用例，分布在 {len(by_file)} 个文件中")
        
        return export_data


def create_sample_config():
    """创建示例配置文件"""
    sample_config = {
            "test_execution_config": {
                "description": "主配置，控制全局测试执行行为。include模式： exclude 模式",
                "execution_mode": "include",
                "test_cases": {
                    "description": "精确指定要执行的测试用例列表，只进行 exclude case 过滤(高优先级，使能之后不再通过 test_filters 过滤case，直接使用精确指定的 case),和 excluded_cases 可用搭配使用",
                    "enabled": False,
                    "cases": []
                },
                "test_filters": {
                    "description": "包含性过滤器，仅执行满足匹配条件的测试用例。和 excluded_tests 搭配使用",
                    "by_file": {
                        "description": "按文件名过滤，仅执行指定文件中的测试。",
                        "enabled": False,
                        "files": []
                    },
                    "by_class": {
                        "description": "按类名过滤，仅执行指定类的测试。",
                        "enabled": False,
                        "classes": []
                    },
                    "by_function": {
                        "description": "按函数名过滤，仅执行指定函数的测试。",
                        "enabled": False,
                        "functions": []
                    },
                    "by_pattern": {
                        "description": "按re包含过滤，仅执行匹配re的测试。",
                        "enabled": False,
                        "patterns": []
                    }
                },
                "excluded_tests": {
                    "description": "模糊排除，按文件、类、函数或模式批量排除测试。",
                    "enabled": False,
                    "by_file": {
                        "description": "按文件名排除测试。",
                        "enabled": False,
                        "files": []
                    },
                    "by_class": {
                        "description": "按类名排除测试。",
                        "enabled": False,
                        "classes": []
                    },
                    "by_function": {
                        "description": "按函数名排除测试。",
                        "enabled": False,
                        "functions": []
                    },
                    "by_pattern": {
                        "description": "按re模式排除测试。",
                        "enabled": False,
                        "patterns": []
                    }
                },
                "excluded_cases": {
                    "description": "精确排除指定的测试用例,只对精确指定的执行case有效。",
                    "enabled": False,
                    "cases": []
                }
            }
        }
    
    with open("test_config_sample.json", "w", encoding="utf-8") as f:
        json.dump(sample_config, f, indent=4, ensure_ascii=False)
    
    print("示例配置文件已创建: test_config_sample.json")


if __name__ == "__main__":
    # 创建示例配置
    create_sample_config()
    
    # 测试过滤器
    filter_tool = TestCaseFilter("test_config_sample.json")
    filter_tool.print_config_info()
    
    # 生成pytest参数示例
    base_args = ["-v", "-s", "Testcases"]
    filtered_args = filter_tool.generate_pytest_args(base_args)
    print(f"\n生成的pytest参数: {filtered_args}")

from ..utils.io import dump_yaml

def emit_observed_yaml(observed_data, output_path):
    """
    将观测数据输出为 YAML 文件
    
    Args:
        observed_data: 观测数据字典
        output_path: 输出文件路径
    """
    dump_yaml(observed_data, output_path)

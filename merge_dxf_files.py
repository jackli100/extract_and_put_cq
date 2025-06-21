import os
import ezdxf
from ezdxf.math import Vec3

def merge_dxf_files(source_directory, output_file):
    """
    将指定目录下的所有DXF文件内容原位粘贴到一个新的DXF文件中
    
    Args:
        source_directory: 源DXF文件目录
        output_file: 输出的合并DXF文件路径
    """
    # 创建新的DXF文档
    merged_doc = ezdxf.new()
    merged_msp = merged_doc.modelspace()
    
    # 用于存储所有已处理的块定义，避免重复
    processed_blocks = set()
    
    print(f"开始合并目录 {source_directory} 下的所有DXF文件...")
    
    dxf_files = [f for f in os.listdir(source_directory) if f.lower().endswith('.dxf')]
    
    if not dxf_files:
        print("未找到任何DXF文件!")
        return
    
    for filename in dxf_files:
        file_path = os.path.join(source_directory, filename)
        print(f"正在处理: {filename}")
        
        try:
            # 读取源DXF文件
            source_doc = ezdxf.readfile(file_path)
            source_msp = source_doc.modelspace()
            
            # 复制块定义（避免重复）
            for block_name, block in source_doc.blocks.items():
                if block_name not in processed_blocks and not block_name.startswith('*'):
                    try:
                        # 在目标文档中创建新块
                        new_block = merged_doc.blocks.new(name=block_name)
                        
                        # 复制块中的所有实体
                        for entity in block:
                            new_block.add_entity(entity.copy())
                        
                        processed_blocks.add(block_name)
                    except ezdxf.DXFValueError:
                        # 块名已存在，跳过
                        pass
            
            # 复制图层定义
            for layer_name, layer in source_doc.layers.items():
                if layer_name not in merged_doc.layers:
                    new_layer = merged_doc.layers.new(name=layer_name)
                    # 复制图层属性
                    new_layer.color = layer.color
                    new_layer.linetype = layer.linetype
                    new_layer.lineweight = layer.lineweight
                    new_layer.transparency = layer.transparency
                    new_layer.on = layer.on
                    new_layer.freeze = layer.freeze
                    new_layer.lock = layer.lock
            
            # 复制线型定义
            for linetype_name, linetype in source_doc.linetypes.items():
                if linetype_name not in merged_doc.linetypes and linetype_name not in ['ByLayer', 'ByBlock']:
                    try:
                        merged_doc.linetypes.new(
                            name=linetype_name,
                            dxfattribs={'description': linetype.dxf.description}
                        )
                    except:
                        pass
            
            # 复制文字样式
            for style_name, style in source_doc.styles.items():
                if style_name not in merged_doc.styles:
                    try:
                        merged_doc.styles.new(
                            name=style_name,
                            dxfattribs={
                                'font': style.dxf.font,
                                'width': style.dxf.width,
                                'height': style.dxf.height,
                                'oblique': style.dxf.oblique
                            }
                        )
                    except:
                        pass
            
            # 复制所有模型空间实体（原位粘贴）
            entity_count = 0
            for entity in source_msp:
                try:
                    # 直接复制实体，保持原始坐标
                    new_entity = entity.copy()
                    merged_msp.add_entity(new_entity)
                    entity_count += 1
                except Exception as e:
                    print(f"  警告: 无法复制实体 {entity.dxftype()}: {e}")
            
            print(f"  成功复制 {entity_count} 个实体")
            
        except Exception as e:
            print(f"  错误: 无法处理文件 {filename}: {e}")
            continue
    
    # 保存合并后的文件
    try:
        merged_doc.saveas(output_file)
        print(f"\n合并完成! 输出文件: {output_file}")
        print(f"总共处理了 {len(dxf_files)} 个DXF文件")
    except Exception as e:
        print(f"保存文件时出错: {e}")

def merge_dxf_with_layers(source_directory, output_file, use_filename_as_layer=True):
    """
    将DXF文件合并，可选择是否将文件名作为图层名
    
    Args:
        source_directory: 源DXF文件目录
        output_file: 输出文件路径
        use_filename_as_layer: 是否将文件名作为图层名
    """
    merged_doc = ezdxf.new()
    merged_msp = merged_doc.modelspace()
    processed_blocks = set()
    
    print(f"开始合并目录 {source_directory} 下的所有DXF文件...")
    
    dxf_files = [f for f in os.listdir(source_directory) if f.lower().endswith('.dxf')]
    
    for filename in dxf_files:
        file_path = os.path.join(source_directory, filename)
        print(f"正在处理: {filename}")
        
        try:
            source_doc = ezdxf.readfile(file_path)
            source_msp = source_doc.modelspace()
            
            # 如果使用文件名作为图层，创建以文件名命名的图层
            layer_name = None
            if use_filename_as_layer:
                layer_name = os.path.splitext(filename)[0]
                if layer_name not in merged_doc.layers:
                    merged_doc.layers.new(name=layer_name)
            
            # 复制必要的定义（块、图层、线型等）
            # ... (同上面的代码)
            
            # 复制实体并设置图层
            entity_count = 0
            for entity in source_msp:
                try:
                    new_entity = entity.copy()
                    if use_filename_as_layer:
                        new_entity.dxf.layer = layer_name
                    merged_msp.add_entity(new_entity)
                    entity_count += 1
                except Exception as e:
                    print(f"  警告: 无法复制实体: {e}")
            
            print(f"  成功复制 {entity_count} 个实体到图层 {layer_name if use_filename_as_layer else '原图层'}")
            
        except Exception as e:
            print(f"  错误: 无法处理文件 {filename}: {e}")
    
    merged_doc.saveas(output_file)
    print(f"\n合并完成! 输出文件: {output_file}")

if __name__ == '__main__':
    # 设置源目录和输出文件
    source_dir = r'C:\Users\Administrator\Documents\沪乍杭补定测\04-任务单\测绘返任务单\dxfs'  # 修改为你的DXF文件目录
    output_path = r'merged.dxf'    # 修改为你想要的输出文件路径
    
    # 方式1: 简单合并，保持原有图层
    merge_dxf_files(source_dir, output_path)
    
    # 方式2: 合并时将每个文件的内容放到以文件名命名的图层中
    # merge_dxf_with_layers(source_dir, output_path, use_filename_as_layer=True)
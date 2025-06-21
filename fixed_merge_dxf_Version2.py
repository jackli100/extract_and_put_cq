import os
import ezdxf

def merge_dxf_files(source_directory, output_file):
    """
    将指定目录下的所有DXF文件内容原位粘贴到一个新的DXF文件中
    修复了 ezdxf 新版本的兼容性问题
    
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
            
            # 复制块定义（修复：使用正确的遍历方法）
            for block in source_doc.blocks:
                block_name = block.name
                if block_name not in processed_blocks and not block_name.startswith('*'):
                    try:
                        # 检查目标文档中是否已存在该块
                        if block_name not in merged_doc.blocks:
                            # 在目标文档中创建新块
                            new_block = merged_doc.blocks.new(name=block_name)
                            
                            # 复制块中的所有实体
                            for entity in block:
                                try:
                                    new_block.add_entity(entity.copy())
                                except Exception as e:
                                    print(f"    警告: 无法复制块实体: {e}")
                            
                            processed_blocks.add(block_name)
                    except Exception as e:
                        print(f"    警告: 无法复制块 {block_name}: {e}")
            
            # 复制图层定义（修复：使用正确的遍历方法）
            for layer in source_doc.layers:
                layer_name = layer.dxf.name
                if layer_name not in merged_doc.layers:
                    try:
                        new_layer = merged_doc.layers.new(name=layer_name)
                        # 复制图层属性
                        if hasattr(layer.dxf, 'color'):
                            new_layer.dxf.color = layer.dxf.color
                        if hasattr(layer.dxf, 'linetype'):
                            new_layer.dxf.linetype = layer.dxf.linetype
                        if hasattr(layer.dxf, 'lineweight'):
                            new_layer.dxf.lineweight = layer.dxf.lineweight
                        if hasattr(layer.dxf, 'flags'):
                            new_layer.dxf.flags = layer.dxf.flags
                    except Exception as e:
                        print(f"    警告: 无法复制图层 {layer_name}: {e}")
            
            # 复制线型定义（修复：使用正确的遍历方法）
            for linetype in source_doc.linetypes:
                linetype_name = linetype.dxf.name
                if (linetype_name not in merged_doc.linetypes and 
                    linetype_name not in ['ByLayer', 'ByBlock', 'Continuous']):
                    try:
                        description = getattr(linetype.dxf, 'description', '')
                        merged_doc.linetypes.new(
                            name=linetype_name,
                            dxfattribs={'description': description}
                        )
                    except Exception as e:
                        print(f"    警告: 无法复制线型 {linetype_name}: {e}")
            
            # 复制文字样式（修复：使用正确的遍历方法）
            for style in source_doc.styles:
                style_name = style.dxf.name
                if style_name not in merged_doc.styles and style_name != 'Standard':
                    try:
                        style_attrs = {}
                        if hasattr(style.dxf, 'font'):
                            style_attrs['font'] = style.dxf.font
                        if hasattr(style.dxf, 'width'):
                            style_attrs['width'] = style.dxf.width
                        if hasattr(style.dxf, 'height'):
                            style_attrs['height'] = style.dxf.height
                        if hasattr(style.dxf, 'oblique'):
                            style_attrs['oblique'] = style.dxf.oblique
                        
                        merged_doc.styles.new(name=style_name, dxfattribs=style_attrs)
                    except Exception as e:
                        print(f"    警告: 无法复制文字样式 {style_name}: {e}")
            
            # 复制所有模型空间实体（原位粘贴）
            entity_count = 0
            for entity in source_msp:
                try:
                    # 直接复制实体，保持原始坐标
                    new_entity = entity.copy()
                    merged_msp.add_entity(new_entity)
                    entity_count += 1
                except Exception as e:
                    print(f"    警告: 无法复制实体 {entity.dxftype()}: {e}")
            
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

def simple_merge_dxf(source_directory, output_file):
    """
    简化版本：只复制实体内容，忽略其他定义
    """
    # 创建新的DXF文档
    merged_doc = ezdxf.new()
    merged_msp = merged_doc.modelspace()
    
    print(f"合并目录: {source_directory}")
    
    # 获取所有DXF文件
    dxf_files = [f for f in os.listdir(source_directory) 
                 if f.lower().endswith('.dxf')]
    
    total_entities = 0
    
    for filename in dxf_files:
        file_path = os.path.join(source_directory, filename)
        print(f"处理文件: {filename}")
        
        try:
            # 读取源文件
            source_doc = ezdxf.readfile(file_path)
            source_msp = source_doc.modelspace()
            
            # 复制所有实体到合并文档（原位粘贴）
            entity_count = 0
            for entity in source_msp:
                try:
                    merged_msp.add_entity(entity.copy())
                    entity_count += 1
                except Exception as e:
                    print(f"    警告: 跳过实体 {entity.dxftype()}: {e}")
            
            total_entities += entity_count
            print(f"  完成: {filename} - 复制了 {entity_count} 个实体")
            
        except Exception as e:
            print(f"  错误: {filename} - {e}")
    
    # 保存合并文件
    try:
        merged_doc.saveas(output_file)
        print(f"\n合并完成! 保存到: {output_file}")
        print(f"总共复制了 {total_entities} 个实体")
    except Exception as e:
        print(f"保存文件时出错: {e}")

if __name__ == '__main__':
    # 设置源目录和输出文件
    source_dir = r'C:\Users\Administrator\Documents\沪乍杭补定测\04-任务单\线路任务单\dxfs'  # 修改为你的DXF文件目录
    output_path = r'提任务单.dxf'    # 修改为你想要的输出文件路径
    
    # 方式1: 简单合并，保持原有图层
    merge_dxf_files(source_dir, output_path)
    
    # 方式2: 合并时将每个文件的内容放到以文件名命名的图层中
    # merge_dxf_with_layers(source_dir, output_path, use_filename_as_layer=True)
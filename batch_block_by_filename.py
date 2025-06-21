import os
import ezdxf

def all_entities_to_block(doc, block_name):
    msp = doc.modelspace()
    entities = list(msp)
    # 如果已经存在同名块，先删除
    if block_name in doc.blocks:
        doc.blocks.remove(block_name)
    block = doc.blocks.new(name=block_name)
    for e in entities:
        block.add_entity(e.copy())
    # 清空原有模型空间
    msp.delete_all_entities()
    # 插入块引用
    msp.add_blockref(block_name, (0, 0))

def process_directory(directory):
    for filename in os.listdir(directory):
        if filename.lower().endswith('.dxf'):
            file_path = os.path.join(directory, filename)
            print(f'Processing {file_path}...')
            doc = ezdxf.readfile(file_path)
            block_name = os.path.splitext(filename)[0]
            all_entities_to_block(doc, block_name)
            doc.saveas(file_path)  # 覆盖原文件
            print(f'Finished {file_path}')

if __name__ == '__main__':
    folder = r'C:\Users\Administrator\Documents\沪乍杭补定测\04-任务单\线路任务单\dxfs'  # 修改为你的目录
    process_directory(folder)
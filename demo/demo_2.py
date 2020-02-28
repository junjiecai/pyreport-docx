from docx import Document

from doc_generator.core import render


def demo_2():
    from shutil import copyfile

    document = Document('template.docx')

    copyfile("template.docx", 'target.docx')

    text_data = {
        'description': "测试描述"
    }
    table_data = {
        "stats": {
            "headers": ["a", "b"],
            "data": [(1, 2), (3, 4)]
        }
    }
    image_data = {
        'image': 'test.png'
    }
    header_data = {
        "title": "Template Demo"
    }
    list_data = {
        "list": ['item_1', 'item_2', 'item_3']
    }
    document_des = Document('target.docx')
    save_path = 'results/desc.docx'

    render(document, document_des, header_data, image_data, list_data, save_path, table_data, text_data)


if __name__ == '__main__':
    demo_2()

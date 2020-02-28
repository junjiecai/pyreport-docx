from doc_generator.core import Doc, OrderedList


def demo_1():
    document = Doc()
    document.add_header('标题1')
    document.add_header('标题1.1', level=2)
    document.add_header('标题1.2', level=2)
    document.add_paragraph("测试报告")
    document.add_ordered_list([
        '项目1',
        '项目2',
        '项目3',
        OrderedList([
            'A',
            'B',
            OrderedList(['aa', 'bb']),
            "C"
        ]
        )
    ]
    )
    document.add_image('test.png')
    document.add_table(["A", "B", "C"], [(1, 2, 3), (2, 3, 4)])
    document.to_docx('results/测试.docx')


if __name__ == '__main__':
    demo_1()

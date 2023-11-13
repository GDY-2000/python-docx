import subprocess

from docxtpl import DocxTemplate

def generate_report(data, template_path='report_template.docx', output_path='report.docx', open_file=True):
    """
    生成报告文档。

    :param data: 包含要在模板中替换的数据的字典。
    :param template_path: 模板文件的路径，默认为'report_template.docx'。
    :param output_path: 生成的文档的路径，默认为'report.docx'。
    :param open_file: 是否在生成文档后打开它，默认为True。
    """
    # 打开模板文件
    doc = DocxTemplate(template_path)
    print(data)
    # 渲染数据到模板
    doc.render(data)

    # 保存渲染后的文档
    doc.save(output_path)

    # 使用subprocess模块打开生成的文档
    try:
        if open_file:
            subprocess.Popen(['start', '', output_path], shell=True)
    except subprocess.CalledProcessError as e:
        print(f'使用subprocess模块打开生成的文档失败: {e}')

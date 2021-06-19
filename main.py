from docx.shared import Mm
from wordmarker.contexts import WordMarkerContext
from wordmarker.templates import CsvTemplate, PdbcTemplate, WordTemplate, AbstractConverter
from docxtpl import InlineImage


class TextConverter(AbstractConverter):
    def __init__(self, word_tpl_: WordTemplate):
        super().__init__(word_tpl_)

    @staticmethod
    def author():
        return 'chensixiang'


if __name__ == '__main__':
    # 初始化上下文
    WordMarkerContext("config.yaml")
    # 读写csv文件的模板
    csv_tpl = CsvTemplate()
    # 读写数据库的模板
    pdbc_tpl = PdbcTemplate()
    # 读写word文档的模板
    word_tpl = WordTemplate()

    # 从配置中获取.csv文件，类型为DataFrame
    csv_dict = csv_tpl.csv_to_df()
    prosperity_index_file = csv_dict['景气指数_加盐.csv']
    print(prosperity_index_file)

    # 将DataFrame类型的数据写入数据库
    pdbc_tpl.update_table(prosperity_index_file, "t_prosperity_index")

    # 从数据库中获取数据，类型为DataFrame
    prosperity_index_database = pdbc_tpl.query_table("t_prosperity_index")
    print(prosperity_index_database)

    # 将DataFrame类型的数据转换为.csv文件
    csv_tpl.df_to_csv({'景气指数_数据库.csv': prosperity_index_database})

    # 创建TextConverter对象，它继承了AbstractConverter，可以将yaml模板中的插值表达式进行转换
    text = TextConverter(word_tpl)

    content = {
        # 直接赋值
        'title': '景气指数',
        # 从example.yaml模板中获取值
        'img_title': word_tpl.get_value("example.img_title"),
        # 图片
        'img': InlineImage(word_tpl.tpl, word_tpl.get_img_file('景气指数.png'),
                           width=Mm(100)),
        # 从example.yaml模板中获取值
        'explanation': word_tpl.get_value("example.explanation"),
        # 从meta.yaml模板中获取将插值表达式进行转换后的值
        'author': text.get_value("meta.author"),
        # 从meta.yaml模板中获取值
        'email': word_tpl.get_value("meta.email"),
    }

    # 添加content到总的上下文中
    word_tpl.append(content)
    # 输出word文档
    word_tpl.build("example.docx")

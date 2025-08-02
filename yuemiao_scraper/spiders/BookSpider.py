import scrapy
import os


class BookSpider(scrapy.Spider):
    name = "book_spider"
    start_urls = ['https://www.quddu.com/book/40679/']  # 初始页面

    def __init__(self):
        # 初始化存储章节链接的列表和输出文件
        self.chapter_list = []
        self.content_dict = {}  # 存储章节内容
        self.output_file = "book_content.txt"
        # 确保输出文件是空的
        if os.path.exists(self.output_file):
            os.remove(self.output_file)

    def parse(self, response):
        # 找到 class="list_dd" 下的所有 a 标签，提取 href 属性
        chapter_links = response.css('.list_dd a::attr(href)').getall()
        self.chapter_list = ['https://www.quddu.com' + link for link in chapter_links]

        # 按顺序请求章节链接
        for index, link in enumerate(self.chapter_list):
            yield scrapy.Request(link, callback=self.parse_chapter, meta={'index': index})

    def parse_chapter(self, response):
        # 提取章节名称
        chapter_title = response.css('.book_con h1::text').get()
        # 提取章节内容，处理替换格式
        raw_content = response.css('#zoom').get()

        formatted_content = (
            raw_content.replace('&nbsp;', '')
            .replace('<br><br><br><br>', '\n')
            .replace('<br>', '\n')
            .replace('ru', '乳')
            .replace('yin水', '淫水')
            .replace('yin蒂', '阴蒂')
            .replace('yin唇', '阴唇')
            .replace('yin道', '阴道')
            .replace('gui头', '龟头')
            .replace('ji巴', '鸡巴')
            .replace('ai', '爱')
            .replace('rou棒', '肉棒')
            .replace('jing液', '精液')
            .replace('高氵朝', '高潮')
            .replace('xiāo穴', '小穴')
            .replace('yin户', '阴户')
            .replace('yin茎', '阴茎')
            .replace('mi穴', '蜜穴')
            .replace('yáng具', '阳具')
            .replace('mi穴', '蜜穴')

        )
        # 使用 Scrapy 的选择器提取纯文本
        chapter_content = scrapy.Selector(text=formatted_content).xpath('//text()').getall()
        chapter_content = ''.join(chapter_content).strip()

        # 存储章节内容到字典，按索引顺序
        index = response.meta['index']
        self.content_dict[index] = f"{chapter_title}\n\n{chapter_content}\n\n"

        # 检查是否所有章节已经爬取完成
        if len(self.content_dict) == len(self.chapter_list):
            self.save_content_to_file()

    def save_content_to_file(self):
        # 按照章节顺序写入文件
        with open(self.output_file, 'a', encoding='utf-8') as f:
            for index in sorted(self.content_dict.keys()):
                f.write(self.content_dict[index])
        self.log("所有章节已保存到文件中")


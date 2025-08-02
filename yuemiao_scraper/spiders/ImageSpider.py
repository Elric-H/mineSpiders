import os
import scrapy
import urllib.parse
from twisted.internet.error import ConnectionLost


class ImageSpider(scrapy.Spider):
    name = "image_spider"

    BROWSER_HEADERS = {
        "connection": "close",
        "referer": "https://xiunice.com/",
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "accept-language": "zh-CN,zh;q=0.9,ja;q=0.8",
        "cookie": "_ga=GA1.1.1412762376.1732553742; _ga_4VHH86F4BG=GS1.1.1739677305.22.1.1739677629.0.0.0",
        "priority": "u=0, i",
        "sec-ch-ua": '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
        "sec-fetch-dest": "document",
        "sec-fetch-mode": "navigate",
        "sec-fetch-site": "none",
        "sec-fetch-user": "?1",
        "upgrade-insecure-requests": "1",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        # "connection": "keep-alive"
    }

    def start_requests(self):
        yield scrapy.Request(
            url="https://xiunice.com/xiuren%e7%a7%80%e4%ba%ba%e7%bd%91-no-5946-%e5%a6%b2%e5%b7%b1_toxic-71p-4k",
            headers=self.BROWSER_HEADERS,
            cookies={  # 单独处理 Cookie
                "_ga": "GA1.1.1412762376.1732553742",
                "_ga_4VHH86F4BG": "GS1.1.1739677305.22.1.1739677629.0.0.0"
            },
            meta={'max_retry_times': 10}
        )

    def download_image(self, response):
        # 保存图片到指定文件夹
        folder_path = response.meta["folder_path"]
        image_name = response.url.split("/")[-1]  # 从URL中提取图片文件名
        image_path = os.path.join(folder_path, image_name)

        # 保存图片到本地
        with open(image_path, "wb") as f:
            f.write(response.body)
        self.log(f"图片已保存: {image_path}")

    def parse(self, response):
        # 获取文件夹名称
        folder_name = response.css("h1.tdb-title-text::text").get()
        if folder_name is None:
            self.logger.warning("未找到文件夹名称，使用默认值")
            folder_name = "default_folder_name"
        folder_path = os.path.join(os.getcwd(), folder_name)

        # 创建文件夹
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # 使用提供的父容器 XPath 提取父容器信息
        parent_container = response.xpath('//*[@id="tdi_78"]/div/div[2]')
        if not parent_container:
            self.logger.error("未找到父容器，请检查 XPath 是否正确")
            return

        # 打印父容器的内容
        self.logger.info(f"父容器内容：{parent_container.get()}")

        # 从父容器中查找所有的 <a> 标签
        links = parent_container.css("img::attr(src)").getall()
        if not links:
            self.logger.error("未从父容器中找到任何链接，请检查选择器")
            return

        self.logger.info(f"找到 {len(links)} 个链接")

        # 下载图片
        for image_link in links:
            yield scrapy.Request(
                image_link,
                callback=self.download_image,
                meta={"folder_path": folder_path},
            )




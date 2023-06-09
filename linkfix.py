import re


def instagram_link_fix(link: str) -> str:
    """
    将Instagram链接统一转换为: "https://www.instagram.com/xxx/" 的格式
    """
    # 在链接最后追加"/"
    link = link.rstrip("/") + "/"

    match = re.search(r'(.*instagram\.com/[\w_/]+)', link)
    if match:
        link = match.group(0).rstrip("/") + "/"
        link_part_user = link.split("instagram")[-1:][0]
        link = "https://www.instagram" + link_part_user

        # 再次检查链接合法性
        if len(link.split("/")) != 5:
            link_parts = link.split("/")

            # 处理"https://www.instagram.com/p/Ckqykz7L2uo/"类型
            if link_parts[3] == "p":
                link = link_parts[:3] + link_parts[4:]
                link = link[0] + "//" + link[2] + "/" + link[3] + "/"

            # 处理"https://wwww.instagram.com/vohr/reels/"类型
            if link_parts[-2] == "reels":
                link = link_parts[0] + "//" + link_parts[2] + "/" + link_parts[3] + "/"

        return link

    # 处理Alexa Hendricks (@achendricks) 窶
    match = re.search(r'\(@.*\)', link)
    if match:
        link = "https://www.instagram.com/" + match.group(0)[2:-2]

    if (re.search(r'[0-9].*(add|Add)', link) or " "in link):
        return ""

def tiktok_link_fix(link: str) -> str:
    """
    将TikTok链接统一转换为: "https://www.tiktok.com/@xxx" 的格式
    """
    pass


if __name__ == "__main__":
    from xlsx_operate import ExcelToPd

    excel_path = "/Users/mokerl/Desktop/总表.xls"

    exceldata = ExcelToPd(excel_path)

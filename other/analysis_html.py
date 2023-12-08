###################################################################################
# 分析“知识星球”网站上的代码，爬取下来
###################################################################################
import os
import time
import traceback
import urllib.request
from bs4 import BeautifulSoup


def get_image(path, save_path):
    while 1:
        try:
            urllib.request.urlretrieve(path, filename=save_path)
        except Exception:
            print(traceback.format_exc())
            print(f"保存图片出错:{path}")
            time.sleep(5)
            continue
        break


if __name__ == '__main__':
    path = r"E:\NewFolder\zhishi\raw.html"
    save_img = r"E:\NewFolder\zhishi\image"
    save_index = 0
    with open(path, "r", encoding="utf-8") as f:
        data = f.read()
    soup = BeautifulSoup(data, features="lxml")
    topic_list = soup.find_all("app-topic")
    print("总共有{}篇".format(len(topic_list)))
    save_txt = os.path.join(os.path.dirname(save_img), "result.txt")
    with open(save_txt, "w", encoding="utf-8") as f:
        for index, tag in enumerate(topic_list):
            last_index = save_index
            total_has_image = tag.find_all("img", "item")
            date = tag.find("div", "date").text
            if len(total_has_image) != 0:
                print("含有图片,图片数量{},索引{},时间{}".format(len(total_has_image), index, date))
            topic_type = None
            content = ""
            talk_content = tag.find("div", "talk-content-container")
            answer_content = tag.find("div", "answer-content-container")
            task = tag.find("div", "task-content-container")
            if answer_content is not None:
                topic_type = "问答"
                question_contain = answer_content.find("div", "question")
                content = "**问题:{}\n".format(question_contain.text)
                has_img = question_contain.find("app-image-gallery")
                if has_img is not None:
                    content += "**图片:"
                    imgs = has_img.find_all("img", "item")
                    for img in imgs:
                        save_path = os.path.join(save_img, "{}.jpg".format(save_index))
                        get_image(img.attrs["src"], save_path)
                        print("保存问题图片:{}".format(save_path))
                        save_index += 1
                        content += os.path.basename(save_path) + ","
                    content += "\n"
                answer = answer_content.find("div", "answer")
                content += "**回答:{}\n".format(answer.text)
                has_img = answer.find("app-image-gallery")
                if has_img is not None:
                    content += "**图片:"
                    imgs = has_img.find_all("img", "item")
                    for img in imgs:
                        save_path = os.path.join(save_img, "{}.jpg".format(save_index))
                        get_image(img.attrs["src"], save_path)
                        print("保存答案图片:{}".format(save_path))
                        save_index += 1
                        content += os.path.basename(save_path) + ","
                    content += "\n"
                single_images = answer_content.find_all("app-image-gallery", recursive=False)
                if len(single_images) != 0:
                    content += "**单独的图片:"
                    for simgs in single_images:
                        imgs = simgs.find_all("img", "item")
                        for img in imgs:
                            save_path = os.path.join(save_img, "{}.jpg".format(save_index))
                            get_image(img.attrs["src"], save_path)
                            print("保存单独的图片:{}".format(save_path))
                            save_index += 1
                            content += os.path.basename(save_path) + ","
                    content += "\n"
            elif talk_content is not None:
                topic_type = "自述"
                content = "**内容:{}\n".format(talk_content.find("div", "content").text)
                has_img = talk_content.find("app-image-gallery")
                if has_img is not None:
                    content += "**图片:"
                    imgs = has_img.find_all("img", "item")
                    for img in imgs:
                        save_path = os.path.join(save_img, "{}.jpg".format(save_index))
                        get_image(img.attrs["src"], save_path)
                        print("保存内容图片:{}".format(save_path))
                        save_index += 1
                        content += os.path.basename(save_path) + ","
                    content += "\n"
            elif task is not None:
                print("作业题目")
                topic_type = "作业"
                content = "**内容;自己找到时间点截图。"
            else:
                print("该项主题存在问题", index)
            save_data = "**时间：{}\n**类型:{}\n{}".format(date, topic_type, content)
            save_data += "-" * 100 + "\n"
            f.write(save_data)
            if topic_type == "作业":
                continue
            assert save_index == last_index + len(total_has_image), index

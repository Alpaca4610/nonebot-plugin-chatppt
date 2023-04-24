import os
import shutil
from pathlib import Path

import nonebot
from nonebot import on_command
from nonebot.adapters.onebot.v11 import MessageEvent, Bot
from nonebot.adapters.onebot.v11 import (MessageSegment)
from nonebot.params import ArgPlainText

from .config import Config
from .core import generate_ppt

plugin_config = Config.parse_obj(nonebot.get_driver().config.dict())

if not plugin_config.slides_limit:
    slides_limit = 10
else:
    slides_limit = plugin_config.slides_limit


def delete_file(dir):
    if os.path.exists(dir):
        shutil.rmtree(dir)


cache_folder = Path() / "data" / "nonebot-plugin-chatppt" / "caches"

delete_request = on_command("删除所有缓存PPT", block=True, priority=1)


@delete_request.handle()
async def _():
    delete_file(cache_folder)
    await delete_request.finish(MessageSegment.text("全部删除文件缓存成功！"), at_sender=True)


delete_user_request = on_command("删除缓存PPT", block=True, priority=1)


@delete_user_request.handle()
async def _(event: MessageEvent):
    delete_file(os.path.join(cache_folder, event.get_session_id()))
    await delete_user_request.finish(MessageSegment.text("成功删除你在该群的缓存文件！"), at_sender=True)


start_request = on_command("chatppt", block=True, priority=1)


@start_request.handle()
async def handle_function():
    directory = './data/nonebot-plugin-chatppt/theme'
    files_string = ''

    for root, dirs, files in os.walk(directory):
        for filename in files:
            files_string += filename + '\n'

    print(files_string)
    await start_request.send(MessageSegment.text(f"当前可用PPT模版有:\n{files_string}"), at_sender=True)


@start_request.got("content", prompt="请输入选择的PPT模版（需要带后缀名）、PPT主题、生成页数。示例：模版：xxxxxx，主题：xxxxx，页数：5")
async def _(bot: Bot, event: MessageEvent, content: str = ArgPlainText()):
    if content == "" or content is None:
        await start_request.finish(MessageSegment.text("内容不能为空！"), at_sender=True)

    theme = content.split("：")[1].split("，")[0]
    topic = content.split("：")[2].split("，")[0]
    length = content.split("：")[3].split("，")[0]
    filepath = './data/nonebot-plugin-chatppt/theme/' + theme

    if not os.path.exists(filepath):
        await start_request.reject(f"PPT模版 {theme} 不存在，请重新输入！")

    if int(slides_limit) < int(length):
        await start_request.finish(MessageSegment.text(f"生成的PPT不能超过{slides_limit}页！"), at_sender=True)

    await start_request.send(MessageSegment.text("生成中......."))

    try:
        res = await generate_ppt(filepath, topic, length, event.get_session_id())
    except Exception as error:
        await start_request.finish(str(error), at_sender=True)

    await bot.upload_group_file(group_id=event.group_id,
                                file=res,
                                name=event.get_session_id() + topic + ".pptx")
    start_request.finish("生成完成！", at_sender=True)

import os
import shutil
from pathlib import Path

import nonebot
from nonebot import on_command, on_notice
from nonebot.adapters.onebot.v11 import (Message, MessageSegment)
from nonebot.adapters.onebot.v11 import MessageEvent, Bot, GroupMessageEvent, GroupUploadNoticeEvent
from nonebot.params import CommandArg

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
delete_file(cache_folder)

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
async def _(bot: Bot, event: MessageEvent, msg: Message = CommandArg()):
    delete_file(os.path.join(cache_folder, event.get_session_id()))

    # if api_key == "":
    #     await start_request.finish(MessageSegment.text("请先配置openai_api_key"), at_sender=True)

    content = msg.extract_plain_text()

    if content == "" or content is None:
        await start_request.finish(MessageSegment.text("内容不能为空！"), at_sender=True)

    topic = content.split("：")[1].split("，")[0]
    length = content.split("：")[2]
    if slides_limit < int(length):
        await start_request.finish(MessageSegment.text(f"生成的PPT不能超过{slides_limit}页！"), at_sender=True)

    await start_request.send(MessageSegment.text("生成中......."))

    try:
        res = await generate_ppt(topic, length, event.get_session_id())
    except Exception as error:
        await start_request.finish(str(error), at_sender=True)

    await bot.upload_group_file(group_id=event.group_id,
                                file=res,
                                name=event.get_session_id() + topic + ".pptx")

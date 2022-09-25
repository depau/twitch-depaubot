#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import asyncio
import os
import shelve
import traceback
from enum import Enum
from typing import cast

import win32com.client as wincom

from xml.sax.saxutils import escape

from twitchio import Message
from twitchio.ext import commands


class SpeechVoiceSpeakFlags(Enum):
    # SpVoice Flags
    SVSFDefault = 0
    SVSFlagsAsync = 1
    SVSFPurgeBeforeSpeak = 2
    SVSFIsFilename = 4
    SVSFIsXML = 8
    SVSFIsNotXML = 16
    SVSFPersistXML = 32

    # Normalizer Flags
    SVSFNLPSpeakPunc = 64

    # Masks
    SVSFNLPMask = 64
    SVSFVoiceMask = 127
    SVSFUnusedFlags = -128


def speak_sync(text: str, lang: str = "en-US"):
    # noinspection PyPep8Naming
    SpVoice = wincom.Dispatch("SAPI.SpVoice")
    SpVoice.Speak(
        f"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='{lang}'>{escape(text)}</speak>",
        SpeechVoiceSpeakFlags.SVSFIsXML.value)


async def speak(text: str, lang: str = "en-US"):
    # noinspection PyPep8Naming
    SpVoice = wincom.Dispatch("SAPI.SpVoice")
    SpVoice.Speak(
        f"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='{lang}'>{escape(text)}</speak>",
        SpeechVoiceSpeakFlags.SVSFIsXML.value | SpeechVoiceSpeakFlags.SVSFlagsAsync.value)
    while not SpVoice.WaitUntilDone(10):
        await asyncio.sleep(0.1)


class Bot(commands.Bot):
    def __init__(self):
        # Initialise our Bot with our access token, prefix and a list of channels to join on boot...
        # prefix can be a callable, which returns a list of strings or a string...
        # initial_channels can also be a callable which returns a list of strings...
        super().__init__(token=os.environ["TWITCH_ACCESS_TOKEN"], prefix='!',
                         initial_channels=[os.environ["TWITCH_CHANNEL"]])
        self.languages = None
        self.queue = {}

        # noinspection PyBroadException
        try:
            with open(os.environ.get("QUEUE_FILE", "at_queue.txt")) as f:
                t = f.read()
                for line in t.split("\n"):
                    if not line.strip():
                        continue
                    line = line.strip()[2:]
                    v, k = list(map(lambda x: x.strip(), line.split(":", 1)))
                    self.queue[k] = v
        except Exception:
            traceback.print_exc()

    def __enter__(self):
        self.languages = shelve.open(os.environ.get("TTS_LANG_PICKLE_PATH", "tts_lang.pickle"))
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback):
        if self.languages:
            self.languages.close()

    def set_user_lang(self, username: str, lang: str):
        if self.languages is None:
            return
        self.languages[username] = lang
        self.languages.sync()

    def get_user_lang(self, username: str):
        if self.languages is None or username not in self.languages:
            return "en-US"
        return self.languages[username]

    async def event_ready(self):
        # Notify us when everything is ready!
        # We are logged in and ready to chat and use commands...
        print(f'Logged in as | {self.nick}')
        print(f'User id is | {self.user_id}')

    async def event_message(self, message):
        # Messages with echo set to True are messages sent by the bot...
        # For now we just want to ignore them...
        if message.echo:
            return

        # Print the contents of our message to console...
        print(message.content)

        if not message.content.startswith(self._prefix) and not message.content.startswith("?"):
            await self.read_message(message)

        # Since we have commands and are overriding the default `event_message`
        # We must let the bot know we want to handle and invoke our commands...
        await self.handle_commands(message)

    async def read_message(self, message: Message, lang=None):
        if not lang:
            lang = self.get_user_lang(message.author.name)

        says = "says"
        user = "User"
        if lang == "it-IT":
            says = "dice"
            user = "L'utente"

        content = message.content
        if content.startswith("!"):
            # Strip command
            content = content.split(" ", 1)[1]

        msg = f"{user} {message.author.name} {says}: {content}"
        if lang == "en-US":
            msg = msg.replace("ddepau", "dee dehp ah hoo")
            msg = msg.replace("depau", "dehp ah hoo")
        elif lang == "it-IT":
            msg = msg.replace("ddepau", "di depau")

        asyncio.get_event_loop().create_task(
            speak(msg, lang=lang)
        )

    @commands.command()
    async def q(self, ctx: commands.Context):
        return

    @commands.command()
    async def ita(self, ctx: commands.Context):
        self.set_user_lang(ctx.author.name, "it-IT")
        await ctx.send(f"{ctx.author.name}, la tua lingua per il TTS è stata impostata all'italiano 🇮🇹🤌")

    @commands.command()
    async def eng(self, ctx: commands.Context):
        self.set_user_lang(ctx.author.name, "en-US")
        await ctx.send(f"{ctx.author.name}, your TTS language has been set to English 🇬🇧🍔")

    @commands.command()
    async def speak(self, ctx: commands.Context):
        await self.read_message(ctx.message, lang="en-US")

    @commands.command()
    async def parla(self, ctx: commands.Context):
        await self.read_message(ctx.message, lang="it-IT")

    @commands.command()
    async def req(self, ctx: commands.Context):
        entry = cast(str, ctx.message.content).replace(f"{self._prefix}req ", "", 1)
        self.queue[entry] = ctx.author.name
        with open(os.environ.get("QUEUE_FILE", "at_queue.txt"), "w") as f:
            for entry, user in self.queue.items():
                f.write(f"- {user}: {entry}\n")
        await ctx.send(f"{ctx.author.name}, your request has been recorded.")

    @commands.command()
    async def queue(self, ctx: commands.Context):
        await ctx.send(str(list(self.queue.keys())))


if __name__ == '__main__':
    # Since IDK how to use Windows
    if os.path.isfile(".env"):
        print("Sourcing environment file")
        with open(".env") as f:
            for line in f.readlines():
                key, val = line.strip().split("=", 1)
                os.environ[key] = val

    with Bot() as bot:
        bot.run()

# https://cafe.naver.com/programgarden/102
import configparser
from slacker import Slacker

class Slack():
    def __init__(self):
        config = configparser.ConfigParser()
        config.read('config.ini')
        self.token = config['FBA']['SLACK_TOKEN']

    def notification(self, pretext=None, title=None, fallback=None, text=None, img_path=None, channel=None, msg_on=True):
        if msg_on:
            attachments_dict = dict()
            attachments_dict['pretext'] = pretext  # test1
            attachments_dict['title'] = title  # test2
            attachments_dict['fallback'] = fallback  # test3
            attachments_dict['text'] = text  # test4

            attachments = [attachments_dict]

            slack = Slacker(self.token)

            slack.chat.post_message(channel=channel, text=None, attachments=attachments, as_user=True)
            # slack.files.upload(img_path, channels='#realtime-msg')
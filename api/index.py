from serverless_wsgi import handle_request
import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from reminder_server import app

def handler(event, context):
    return handle_request(app, event, context)

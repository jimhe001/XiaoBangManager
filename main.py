#!/usr/bin/python
# -*- coding: utf-8 -*
import ConfigParser
import os
import sys

import tornado.httpserver
from tornado.options import define, options
import tornado.web
import torndb
import base.session


from agent.AgentHandler import  AgentHandler
from agent.HomeHandler import AgentHomeHandler
from agent.IndexHandler import AgentIndexHandler
from agent.LoginHandler import AgentLoginHandler
from agent.LoginHandler import AgentLogoutHandler
from bankService.bankService import BankServiceHandler
from bankService.edit import EditorHandler
from bankService.exportExcel import ExportBankServiceHandler
from bankService.upload import UploadHandler
from base.HomeHandler import HomeHandler
from base.Login import LoginHandler
from base.Login import LogoutHandler
from base.Welcome import WelcomeHandler
from company.Company import CompanyHandler
from customer.CustomerManagerHandler import CustomerManagerHandler
from customer.ExportCustomerHandler import ExportCustomerHandler
from sales.SalesStatistics import SalesStatisticsHandler
from leaflet.Leaflet import LeafletHandler
from sales.ExportSalesHandler import ExportSalesHandler
from version.Version import VersionHandler



#读取配置文件
st = ConfigParser.ConfigParser()
st.read("conf/Db.conf")

#图片路径
STATIC_PATH = os.path.join(os.path.dirname(__file__), "static")
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "templates")

#数据库链接
define("port", default=st.get("db", "binding_port"), help="run port", type=int)
define("mysql_host", default=st.get("db", "db_host"), help="db host")
define("mysql_database", default=st.get("db", "db_database"), help="db name")
define("mysql_user", default=st.get("db", "db_user"), help="db user")
define("mysql_password", default=st.get("db", "db_pass"), help="db password")

def login_required(f):   
    def _wrapper(self,*args, **kwargs):   
        print self.get_current_user()   
        logged = self.get_current_user()   
        if logged == None:   
            self.write('no login')   
            self.finish()   
        else:   
            ret = f(self,*args, **kwargs)   
    return _wrapper   

class Application(tornado.web.Application):
    def __init__(self):
        handlers = [
        (r'/index', WelcomeHandler),
        (r'/home', HomeHandler),
        (r'/', IdexHandler),
        (r'/login', LoginHandler),
        (r'/logout', LogoutHandler),
        (r'/company/(.*)', CompanyHandler),
        (r'/version', VersionHandler),
        (r'/bankService/(.*)', BankServiceHandler),
        (r'/exportBankService/(.*)',ExportBankServiceHandler),
        (r'/exportCustomer/(.*)',ExportCustomerHandler),
        (r'/customer/(.*)',CustomerManagerHandler),
        (r'/exportSales/(.*)',ExportSalesHandler),
        (r'/agent/(.*)',AgentHandler),
        (r'/editorBankService/(.*)',EditorHandler),
        (r'/upload',UploadHandler),
        (r'/agentLogin',AgentLoginHandler),
        (r'/agentLoginOut',AgentLogoutHandler),
        (r'/agentIndex',AgentIndexHandler),
        (r'/agentHome',AgentHomeHandler),
        (r'/leafletService/(.*)',LeafletHandler),
        (r'/SalesStatistics/(.*)',SalesStatisticsHandler)
        
                
        ]
        settings = {
            "autoescape" : None,       
            "xsrf_cookies": True,  #XSRF保护   防止跨站攻击
            "login_url": '/login',
            "cookie_secret": "bXZ/gDAbQA+zaTxdqJwxKa8OZTbuZE/ok3doaow9N4Q=",            
            "template_path":TEMPLATE_PATH,
            "static_path":STATIC_PATH,
            "debug":True,
            "store_options" : {'redis_host': '10.194.0.251','redis_port': 6379, 'redis_pass': ''},
            "session_secret" : "3cdcb1f00803b6e78ab50b466a40b9977db396840c28307f428b25e2277f1bcc",
            "session_timeout" : 60     
        }
        
        tornado.web.Application.__init__(self, handlers, **settings)
        self.session_manager = base.session.SessionManager(settings["session_secret"], settings["store_options"], settings["session_timeout"])
        self.db = torndb.Connection(
            host=options.mysql_host,
            database=options.mysql_database,
            user=options.mysql_user,
            password=options.mysql_password,
            time_zone="+8:00"
        )

class IdexHandler(tornado.web.RequestHandler):
    def get(self):
        self.render("salesStatistics/index.html")

               
def main():
    '''
 Description : 主函数
 
 Input : None

 Output : None
    '''
#     tornado.options.parse_command_line() 
    app = tornado.httpserver.HTTPServer(Application()) 
    app.listen(options.port)
#     app.listen(sys.argv[1].split("=")[1])
    tornado.ioloop.IOLoop.instance().start() 



    
#运行进程
if __name__ == "__main__":
    main()
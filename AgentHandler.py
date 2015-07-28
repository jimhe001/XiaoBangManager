# coding:utf-8
#!/usr/bin/python
'''
==============
  程序：代理商管理
   版本：0.5  
   作者：王树涛
   日期：2015-06-01
   语言：Python 2.7  
   功能：用于登陆
==============
'''
from base.BaseHandler import BaseHandler
import xml.etree.cElementTree as ET
import xlwt
import StringIO
from tornado import web
import traceback
import os
import random
from xlrd import open_workbook
import tornado
import hashlib
tree = ET.ElementTree(file="conf/AgentSql.xml")
# 图片路径
STATIC_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "static")
TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates")

class AgentHandler(BaseHandler):
    
    @tornado.web.authenticated
    def get(self, action):
        if action == "index":
            self.index()
        elif action=="view":
            self.view()
        elif action=="toUpdate":
            self.toUpdate()    
        elif action == "search":
            self.search();
             
        elif action == "exportXsl": 
            self.exportXsl()
        elif action == "toAdd":
            self.toAdd()
        elif action == "toView":
            self.toView()
        elif action == "update":
            self.toUpdate()  
        elif action == "v":
            self.validate()               
        pass
    
    def toUpdate(self):
        db=self.application.db
        agentId=self.get_argument("tid");
        sql1 = self.getSqlById("getUserByCompaynId");
        sql = self.getSqlById("getAgentInfoById");
        uid = self.get_secure_cookie("userId");
        userList = db.query(sql1, uid)
        agentInfo=db.query(sql,agentId)
        agent=agentInfo[0]
        queryAgentDepositRecordList=db.query(self.getSqlById("queryAgentDepositRecordList"),agentId)
        queryAgentPayRecrodList=db.query(self.getSqlById("queryAgentPayRecrodList"),agentId)
        self.render("agent/update.html",agent=agent,userList=userList,adrList=queryAgentDepositRecordList,aprList=queryAgentPayRecrodList)  
    
    def view(self):
        db=self.application.db
        agentId=self.get_argument("tid");
        sql1 = self.getSqlById("getUserByCompaynId");
        sql = self.getSqlById("getAgentInfoById");
        uid = self.get_secure_cookie("userId");
        userList = db.query(sql1, uid)
        agentInfo=db.query(sql,agentId)
        agent=agentInfo[0]
        queryAgentDepositRecordList=db.query(self.getSqlById("queryAgentDepositRecordList"),agentId)
        queryAgentPayRecrodList=db.query(self.getSqlById("queryAgentPayRecrodList"),agentId)
        self.render("agent/view.html",agent=agent,userList=userList,adrList=queryAgentDepositRecordList,aprList=queryAgentPayRecrodList)  
        pass
    @tornado.web.authenticated
    def post(self, action):
        # 查询
        if action == "doAdd":
            self.doAdd();
        if action=="doUpdate":
            self.doUpdate()
        if action == "search":
            self.search()
        elif action == "exportXsl": 
            self.exportXsl()    
        elif action == "upload": 
            self.upload()     
        elif action=="deleteAll":
            self.deleteAll()       
        pass
    
    # 添加视图
    def toAdd(self):
        db = self.application.db;
        sql = self.getSqlById("getUserByCompaynId");
        uid = self.get_secure_cookie("userId");
        userList = db.query(sql, uid)
        self.render("agent/add.html", userList=userList)   
        pass
    # 进行入库操作
    def doAdd(self):
        db = self.application.db;
        try:
            uid = self.get_secure_cookie("userId");
            companyIdInfos = db.query(self.getSqlById("queryCompanyId"),uid)
            companyId=companyIdInfos[0]["company_id"]
            print companyId
            agent_name = self.get_argument("agent_name");
            username=self.get_argument("username")
            password=self.get_argument("password")
            #print password
            m=hashlib.md5()
            m.update(password)
            password=m.hexdigest()
            #print password
            link_name = self.get_argument("link_name");
            link_phone = self.get_argument("link_phone");
            address = self.get_argument("address");
            legal_person_name = self.get_argument("legal_person_name");
            legal_person_idcard_no = self.get_argument("legal_person_idcard_no");
            legal_person_level = self.get_argument("legal_person_level");
            area = self.get_argument("area");
            status = self.get_argument("status");
            discount_rate = self.get_argument("discount_rate");
            avaliabe_money = self.get_argument("avaliabe_money");
            agentInfoSql = self.getSqlById("addAgentInfo");
            agentId=db.execute(agentInfoSql, agent_name, link_name, link_phone, address, legal_person_name, legal_person_idcard_no, legal_person_level, area, status, discount_rate, avaliabe_money, uid, companyId,username,password);
            if agentId>0:
                self.addDepositRecord(db,agentId, uid, companyId)
                self.addAgentPayRecrod(db,agentId, uid, companyId)
            #db.commit();
            self.redirect("index")    
        except Exception:
            #db.rollback()
            traceback.print_exc()
             
    
    ###########################
    '''
              编辑
    '''
    def doUpdate(self):
        #编辑主表信息
        db = self.application.db;
        try:
            uid = self.get_secure_cookie("userId");
            agentId=self.get_argument("agentId")
            companyIdInfos = db.query(self.getSqlById("queryCompanyId"),uid)
            companyId=companyIdInfos[0]["company_id"]
            agent_name = self.get_argument("agent_name");
            link_name = self.get_argument("link_name");
            link_phone = self.get_argument("link_phone");
            address = self.get_argument("address");
            legal_person_name = self.get_argument("legal_person_name");
            legal_person_idcard_no = self.get_argument("legal_person_idcard_no");
            legal_person_level = self.get_argument("legal_person_level");
            area = self.get_argument("area");
            status = self.get_argument("status");
            discount_rate = self.get_argument("discount_rate");
            avaliabe_money = self.get_argument("avaliabe_money");
            updateAgentInfoSql = self.getSqlById("updateAgentInfo");
            #print updateAgentInfoSql
            #print avaliabe_money
            db.execute(updateAgentInfoSql, agent_name, link_name,avaliabe_money, link_phone, address, legal_person_name, legal_person_idcard_no, legal_person_level, area, status, discount_rate, agentId);
            #更新子表
            self.updateDepositRecord(db,agentId, uid, companyId)
            self.updateAgentPayRecrod(db,agentId, uid, companyId)
            #db.commit();
            self.redirect("index")    
        except Exception:
            #db.rollback()
            traceback.print_exc() 
        pass
    
    #更新保证金交纳记录
    def updateDepositRecord(self,db,agentId,userId,companyId):
        #主键
        tids=self.get_arguments("tab1_tid")
        #接收人id
        reciverIds=self.get_arguments("tab1_input_id")
        #缴纳保证金
        payMoneys=self.get_arguments("tab1_money");
        #缴纳时间
        times=self.get_arguments("tab1_time");
        updateSql=self.getSqlById("updateAgentDepositRecrod");
        #print updateSql
        addSql=self.getSqlById("addDepositRecord");
        #先发一条sql 吧dele_falg置为0
        deleeSql=self.getSqlById("deleteAgentDepositRecrod");
        db.execute(deleeSql,agentId)
        #print addSql
        if len(tids)>0:
            for i in range(len(tids)):
                tid=tids[i]
                reciverId=reciverIds[i]
                money=payMoneys[i]
                time=times[i]
                if int(tid)>0:
                    db.execute(updateSql,reciverId,money,time,1,int(tid))
                else:
                    if int(reciverId)>0:
                        #新增的数据
                        db.execute(addSql,reciverId,money,time,agentId,userId,companyId)
    
    def validate(self):
        self.render("agent/demo.html")
    
    
    #代理商预付款交纳记录
    def  updateAgentPayRecrod(self,db,agentId,userId,companyId):
        #主键
        tids=self.get_arguments("tab2_tid")
        #接收人id
        reciverIds=self.get_arguments("tab2_input_id")
        #缴费人
        tab2_jnnames=self.get_arguments("tab2_jnname");
        #缴纳时间
        times=self.get_arguments("tab2_time");
        updateSql=self.getSqlById("updateAgentPayRecrod");
        #print updateSql
        addSql=self.getSqlById("addAgentPayRecrod")
        #先发一条sql 吧dele_falg置为0
        deleeSql=self.getSqlById("deleteAgentPayRecrod");
        db.execute(deleeSql,agentId)
        if len(tids)>0:
            for i in range(len(tids)):
                tid=tids[i]
                reciverId=reciverIds[i]
                jnname=tab2_jnnames[i]
                time=times[i]
                if int(tid)>0:
                    db.execute(updateSql,reciverId,jnname,time,1,tid)
                else:
                    if int(reciverId)>0:
                        #新增的数据
                        db.execute(addSql,reciverId,jnname,time,agentId,userId,companyId)
    pass                
    ########################
    '''
                   添加保证金交纳记录 
    '''      
    def addDepositRecord(self,db,agentId,userId,companyId):
        try:
            sql=self.getSqlById("addDepositRecord")
            #接收人
            receiveIds=self.get_arguments("tab1_input_id");
            #缴纳保证金
            payMoneys=self.get_arguments("tab1_money");
            #缴纳时间
            times=self.get_arguments("tab1_time");
            if len(receiveIds)>0:
                for i in range(len(receiveIds)):
                    receiveId=receiveIds[i]
                    if  receiveIds!="":
                        money=payMoneys[i]
                        time=times[i]
                        db.execute(sql,receiveId,money,time,agentId,userId,companyId)
        except Exception:
            traceback.print_exc()
            self.write("添加保证金交纳记录失败")
            raise Exception("添加保证金交纳记录失败")

    #添加代理商预付款交纳记录    
    def addAgentPayRecrod(self,db,agentId,userId,companyId):
        try:
            db=self.application.db 
            sql=self.getSqlById("addAgentPayRecrod")
            #接收人
            receiveIds=self.get_arguments("tab2_input_id");
            tab2_jnname=self.get_arguments("tab2_jnname")
            #缴纳时间
            times=self.get_arguments("tab2_time");
            if len(receiveIds)>0:
                for i in range(len(receiveIds)):
                    receiveId=receiveIds[i]
                    if  receiveIds!="":
                        jnname=tab2_jnname[i]
                        time=times[i]
                        db.execute(sql,receiveId,jnname,time,agentId,userId,companyId)
        except Exception:
            traceback.print_exc()
            raise Exception("添加保证金交纳记录失败")
    ##############################################################
    def index(self):
        try:
            # 获取用户的id
            uid = self.get_secure_cookie("userId");
            db = self.application.db
            sql = self.getSqlById("getAgentList")
            resultList = db.query(sql, uid)
            self.render("agent/agentlist.html",params=resultList)     
        except Exception:
            traceback.print_exc() 
            pass
    #
    # 导出数据
    #
    def exportXsl(self):
        try:
            customerType = self.get_argument("customerType", 0)
            startTime = self.get_argument("startTime", '')
            endTime = self.get_argument("endTime", '')
            web.RequestHandler.set_header(self, "Content-Disposition", "attachment;filename=customer.xls")
            wherePart = ''
            # 动态拼接查询参数
            if int(customerType) != 0:
                wherePart = wherePart + ' and cu.customer_flag=' + customerType
            if len(startTime) > 0:
                wherePart = wherePart + " and cu.create_time>=" + "'" + startTime + " 00:00:00' "
            if len(endTime) > 0:
                wherePart = wherePart + ' and cu.create_time<=' + "'" + endTime + " 00:00:00' "
            print "wherePart:" + wherePart
            db = self.application.db
            # 获取sql id
            sql = self.getSqlById("searchCustomerList")
            if len(wherePart) > 0:
                sql = sql.replace("@", wherePart)
            else:
                sql = sql.replace("@", " ")
            print "sql:" + sql
            uid = self.get_secure_cookie("userId")
            resultList = db.query(sql, uid)
            # 创建xsl
            workbook = xlwt.Workbook(encoding='utf-8')
            worksheet = workbook.add_sheet('客户数据统计')
            style = xlwt.XFStyle()  # 初始化样式
            alignment = xlwt.Alignment()
            # 水平居中
            alignment.horz = xlwt.Alignment.HORZ_CENTER
            alignment.vert = xlwt.Alignment.VERT_CENTER
            style.alignment = alignment
            # 为样式创建字体
            font = xlwt.Font()  
            # font.name = name # 'Times New Roman'
            font.bold = True
            font.color_index = 4
            style.font = font
            # 创建标题
            worksheet.write_merge(0, 1, 0, 11, '客户数据统计', style)
            # 创建表头
            worksheet.write(2, 0, label="客户名称")
            worksheet.write(2, 1, label='所属企业')
            worksheet.write(2, 2, label='客户标识')
            worksheet.write(2, 3, label='客户类型')
            worksheet.write(2, 4, label='客户级别')
            worksheet.write(2, 5, label='客户规模')
            worksheet.write(2, 6, label='主联系人')
            worksheet.write(2, 7, label='主联方式')
            worksheet.write(2, 8, label='地址')
            worksheet.write(2, 9, label='客户跟进人')
            worksheet.write(2, 10, label='客户创建时间')
            
            row = 3
            # 写入数据
            for i in range(len(resultList)):
                # print "i"+i+" list:"+resultList[i]
                worksheet.write(row, 0, resultList[i]['name'].encode('utf-8'))
                worksheet.write(row, 1, resultList[i]['cpname'].encode('utf-8'))
                worksheet.write(row, 2, resultList[i]['customer_flag'].encode('utf-8'))
                worksheet.write(row, 3, resultList[i]['customer_type'].encode('utf-8'))
                worksheet.write(row, 4, resultList[i]['level'].encode('utf-8'))
                worksheet.write(row, 5, resultList[i]['scale'].encode('utf-8'))
                worksheet.write(row, 6, resultList[i]['main_linkman_name'].encode('utf-8'))
                worksheet.write(row, 7, resultList[i]['main_linkman_phone'].encode('utf-8'))
                worksheet.write(row, 8, resultList[i]['address'].encode('utf-8'))
                worksheet.write(row, 9, resultList[i]['manger_name'].encode('utf-8'))
                worksheet.write(row, 10, str(resultList[i]['create_time']))
                row = row + 1
            sio = StringIO.StringIO()
            workbook.save(sio)
            self.write(sio.getvalue())
            self.finish()
        except Exception:
            self.finish("查询失败")
            traceback.print_exc()     
        pass
    
    
    ##############################
    
    # 读取数据
    def readXsl(self):
        pass 
    
    # 搜索
    def search(self):
        try:
            customerType = self.get_argument("customerType", 0)
            startTime = self.get_argument("startTime", '')
            endTime = self.get_argument("endTime", '')
            wherePart = ''
            # 动态拼接查询参数
            if int(customerType) != 0:
                wherePart = wherePart + ' and cu.customer_flag=' + customerType
            if len(startTime) > 0:
                wherePart = wherePart + " and cu.create_time>=" + "'" + startTime + " 00:00:00' "
            if len(endTime) > 0:
                wherePart = wherePart + ' and cu.create_time<=' + "'" + endTime + " 00:00:00' "
            print "wherePart:" + wherePart
            db = self.application.db
            # 获取sql id
            sql = self.getSqlById("searchCustomerList")
            if len(wherePart) > 0:
                sql = sql.replace("@", wherePart)
            else:
                sql = sql.replace("@", " ")
            print "sql:" + sql
            uid = self.get_secure_cookie("userId")
            resultList = db.query(sql, uid)
            # resultList=json.dumps(resultList,ensure_ascii=False,separators=(',', ':'),cls=CJsonEncoder)
            html = self.createHtmlCotent(resultList)
            self.finish(html)
        except Exception:
            self.finish("查询失败")
            traceback.print_exc()   
    
    # 生成html片段返回浏览器
    def createHtmlCotent(self, resultList):
        if len(resultList) < 0:
            return ''    
        else:
            html = '<thead><tr><th >客户名称</th><th >所属企业</th><th >客户标识</th><th >客户类型</th><th >客户级别</th><th >客户规模</th><th >主联系人</th><th >主联方式</th><th >地址</th><th >客户跟进人</th><th >客户创建时间</th></tr></thead>'
            html = html + '<tbody>'
            for i in range(len(resultList)):
                html = html + '<tr>'
                html = html + '<td>' + resultList[i]['name'].encode('utf-8') + '</td>'
                html = html + '<td>' + resultList[i]['cpname'].encode('utf-8') + '</td>'
                html = html + '<td>' + resultList[i]['customer_flag'].encode('utf-8') + '</td>'
                html = html + '<td>' + resultList[i]['customer_type'].encode('utf-8') + '</td>'
                html = html + '<td>' + resultList[i]['level'].encode('utf-8') + '</td>'
                html = html + '<td>' + resultList[i]['scale'].encode('utf-8') + '</td>'
                html = html + '<td>' + resultList[i]['main_linkman_name'].encode('utf-8') + '</td>'
                html = html + '<td>' + resultList[i]['main_linkman_phone'].encode('utf-8') + '</td>'
                html = html + '<td>' + resultList[i]['address'].encode('utf-8') + '</td>'
                html = html + '<td>' + resultList[i]['manger_name'].encode('utf-8') + '</td>'
                html = html + '<td>' + str(resultList[i]['create_time']) + '</td>'
                html = html + '</tr>'
            html = html + '</tbody>'
            print html
        return html   
        pass
    #############获取sql语句#######################        
    def getSqlById(self, id_):
        idTemp = 'branch[@name="' + id_ + '"]'
        for elem in tree.iterfind(idTemp):
            sql = elem.text
        return sql
    
    
    # 上传文件
    def upload(self):
        try:
            upload_path = os.path.join(STATIC_PATH, '/customer')  # 文件的暂存路径
            # print upload_path
            # 文件目录不存在则创建文件
            if os.path.exists(upload_path) == False:
                os.makedirs(upload_path)
            file_metas = self.request.files['filename'][0]  # 提取表单中‘name’为‘file’的文件元数据
            imageid = str(random.randint(100000, 1000000))
            filename = imageid + file_metas['filename']
            filePath = os.path.join(upload_path, filename)
            print filePath
            with open(filePath, 'wb') as up:  # 有些文件需要已二进制的形式存储，实际中可以更改
                up.write(file_metas['body'])
                pass
            # 打开xls文件开始解析
            self.parseExcel(filePath);
            # self.write("上传成功")
            self.redirect("index")
        except Exception:
            self.write("上传文件失败请检查文件格式是否正确")
        finally:
            # 删除上传的文件
            os.remove(filePath)
               
        
        
    #解析excel 文件 ############################
    def parseExcel(self, filePath):
        db = self.application.db
        queryManagerSql = self.getSqlById("queryManagerId")
        userId = self.get_secure_cookie("userId")
        companyId = db.query(self.getSqlById("queryCompanyId"), userId)
        # 添加客户信息
        customerSql = self.getSqlById("addCustomerInfo")
        try:
            wb = open_workbook(filePath)
            # 得到第一张表单，两种方式：索引和名字    
            sh = wb.sheet_by_index(0)
            for rownum in range(3, sh.nrows):
                print rownum
                rowValueList = sh.row_values(rownum)
                # 客户标识
                customer_flag = rowValueList[2]
                customer_type = rowValueList[3]
                # 客户标识 (1企业客户,2个人客户)
                if  customer_flag.encode('utf-8') == "个人客户":
                    rowValueList[1] = companyId[0]["company_id"]
                    rowValueList[2] = 2 
                elif customer_type.encode('utf-8') == "企业客户":
                    rowValueList[1] = companyId[0]["company_id"]
                    rowValueList[2] = 1
                else:
                    rowValueList[1] = companyId[0]["company_id"]
                    rowValueList[2] = 2   
                # 客户类型(1潜在,2意向,3成交客户)    
                if customer_type.encode('utf-8') == "潜在客户":
                    rowValueList[3] = 1
                elif customer_type.encode('utf-8') == "意向客户":
                    rowValueList[3] = 2
                elif customer_type.encode('utf-8') == "成交客户": 
                    rowValueList[3] = 3
                else:
                    rowValueList[3] = -1     
                # 查询客户跟进人id
                mangerName = rowValueList[9]
                print queryManagerSql
                manger_idList = db.query(queryManagerSql, mangerName.encode('utf-8'), userId)
                if len(manger_idList) > 0:
                    manger_id = manger_idList[0]["tid"]
                else:
                    manger_id = 0    
                print manger_id
                # 默认不关联企业id 设置为0
                parent_id = 0
                result = db.execute(customerSql, rowValueList[1], parent_id, rowValueList[0], rowValueList[2], rowValueList[3], rowValueList[4], rowValueList[5], rowValueList[6], rowValueList[7], rowValueList[8], manger_id, userId, 0)
                print result
        except Exception:
            traceback.print_exc() 
            raise Exception('数据导入失败') 
        pass
    
    ##删除代理商信息
    def deleteAll(self):
        try:
            db=self.application.db;
            sql=self.getSqlById("deleteAgentInfo");
            print sql
            ids=self.get_argument("ids");
            idList=ids.split(":")
            if len(idList)>0:
                for i in range(len(idList)):
                    tid=idList[i]
                    db.execute(sql,tid)
            self.write("ok")            
        except Exception:
            self.write("error")
            traceback.print_exc() 
            
            
            
            
             

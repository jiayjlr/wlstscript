# -*- coding: utf-8 -*-
import  sys
import 	os
import	ConfigParser
import 	xlsxwriter
import 	time
import 	wl
from	xlsxwriter.utility import xl_range
import 	json
from 	collections import OrderedDict

checkfile = "wls-daily-check"

def connectDomain(username,password,adminurl):
    print "\n**********************************************************************"
    print "* Beginning to connect " + adminurl + " with user : " + username
    print "**********************************************************************"
    wl.connect(username, password, adminurl)


def disDomainConnect():
    print "\n**********************************************************************"
    print "* Disconnect the connection ....."
    print "**********************************************************************"
    wl.disconnect()


def exitWLST():
    print "\n**********************************************************************"
    print "* Exit the wlst command line ....."
    print "**********************************************************************"
    sys.exit()


#根据传入的code返回server的状态
def getServerHealthStateByCodeNum(code):
    return {
        0 : 'OK',
        1 : 'WARNING',
        2 : 'CRITICAL',
        3 : 'FAILED',
        4 : 'OVERLOADED',
    }.get(code, "N/A")

#根据传入的url获取服务ip地址和端口号
def getListenAddressPort(wlserver_default_url):
	default_url_cut_start = wlserver_default_url.index('//') + 2
	default_url_cut_end = wlserver_default_url.rindex(':')
	wlserver_listen_address = wlserver_default_url[default_url_cut_start:default_url_cut_end]
	listenPort = wlserver_default_url[default_url_cut_end+1:]
	return wlserver_listen_address,listenPort
	
	
def writeXls(workbook,domainName,username,password,adminurl):

	'''
	worksheet.add_table('A9:G20', {'columns': [{'header': u'序号'},{'header': u'Server名称'},{'header': u'Server监听地址'},{'header': u'Server监听端口'},{'header': u'Cluster'},{'header': u'Server启动状态'},{'header': u'Server健康状态'}]})
	'''
	
	
	
	connectDomain(username,password,adminurl)
	servers = wl.domainConfig().getServers()
	#print 'server count is : ',len(servers) ##输出的结果如果可以的话，从新定义  worksheet.add_table('A9:F20')
	
	
	worksheet = workbook.add_worksheet(domainName)
	workformat = workbook.add_format()
	workformat.set_bold()
	workformat.set_font_color('red')
	workformat.set_font_size(20)
	workformat.set_align('center')
	workformat.set_align('vcenter')
	workformat.set_locked(True)
	workformat.set_border(2)
	worksheet.merge_range(xl_range(0, 0, 4, 10), u'WebLogic域('+domainName+u')中配置和运行时信息', workformat)
	ItemStyle = workbook.add_format({
        'font_size':10,                #字体大小
        'bold':True,                   #是否粗体
        'align':'justify',             #居中对齐
        'top':1,                       #上边框 
        'left':1,                      #左边框
        'right':1,                     #右边框
        'bottom':1                     #底边框
	})
	
	worksheet.set_column(0,1,8,ItemStyle)
	worksheet.set_column(1,10,20,ItemStyle)
	
	worksheet.write('A6', u'操作人',ItemStyle)
	worksheet.write('B6','admin',ItemStyle)
	worksheet.write('C6',u'操作时间',ItemStyle)
	worksheet.write('D6',time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())),ItemStyle)
	worksheet.write('E6',u'报告生成时间',ItemStyle)
	
	
	domainformat = workbook.add_format()
	domainformat.set_bold()
	domainformat.set_font_color('black')
	domainformat.set_font_size(20)
	domainformat.set_align('justify')
	domainformat.set_locked(True)
	domainformat.set_border(1)
	
	
	
	i = 1
	j = 10
	worksheet.merge_range(xl_range(6, 0, 7, 6), u'WebLogic基本配置及状态信息', domainformat)
	worksheet.add_table('A9:G'+bytes(len(servers) + j), {'columns': [{'header': u'序号'},{'header': u'Server名称'},{'header': u'Server监听地址'},{'header': u'Server监听端口'},{'header': u'Cluster集群'},{'header': u'Server启动状态'},{'header': u'Server健康状态'}]})
	
	
	try:
		### get server status 
		for server in servers:
			server_name = server.getName()
			cluster_mbean = server.getCluster()
			mbean_server = wl.getMBean('domainRuntime:/ServerRuntimes/' + server_name)
			if	mbean_server:
				#server_host_url = mbean_server.getListenAddress() ###从新修改
				server_host_url = mbean_server.getIPv4URL('t3')
				'''
				default_url_cut_start = server_host_url.index('//') + 2
				default_url_cut_end = server_host_url.rindex(':')
				wlserver_listen_address = server_host_url[default_url_cut_start:default_url_cut_end]	
				server_port = mbean_server.getListenPort()
				'''
				(wlserver_listen_address,server_port) = getListenAddressPort(server_host_url)
				server_state = mbean_server.getState()
				server_HealthState = mbean_server.getHealthState()
				if 	cluster_mbean:
					cluster_name = cluster_mbean.getName()
					data = [i,server_name,wlserver_listen_address,server_port,cluster_name,server_state,getServerHealthStateByCodeNum(server_HealthState.getState())] #验证server_HealthState
					worksheet.write_row('A'+bytes(j), data)
				else:
					data = [i,server_name,wlserver_listen_address,server_port,'N/A',server_state,getServerHealthStateByCodeNum(server_HealthState.getState())] #验证server_HealthState
					worksheet.write_row('A'+bytes(j), data)
				
			else:
				unknowbean_server = wl.getMBean('domainConfig:/Servers/' + server_name)
				server_host = unknowbean_server.getListenAddress()
				server_port = unknowbean_server.getListenPort()
				server_state = 'UNKNOWN'
				server_HealthState = 'N/A'				
				if 	cluster_mbean:
					cluster_name = cluster_mbean.getName()
					data = [i,server_name,server_host,server_port,cluster_name,server_state,server_HealthState] #验证server_HealthState
					worksheet.write_row('A'+bytes(j), data)
				else:
					data = [i,server_name,server_host,server_port,'N/A',server_state,server_HealthState] #验证server_HealthState
					worksheet.write_row('A'+bytes(j), data)					

				if len(server_host) == 0:
					worksheet.write_comment('C'+bytes(j), ' server ip not config ')
				worksheet.write_comment('G'+bytes(j), ' server not running ')
			i += 1
			j += 1

		
									  
		#'A'+bytes(j)+':'+'F'+bytes(l)
		worksheet.merge_range(xl_range(j+2, 0, j+3, 5), u'WebLogic执行线程使用信息', domainformat)
		worksheet.add_table('A'+bytes(j+5)+':F'+bytes(len(servers) + j+6), {'columns': [{'header': u'序号'},{'header': u'Server名称'},{'header': u'活动的执行线程数'},{'header': u'执行线程繁忙率 '},{'header': u'Hogging线程数'},{'header': u'Stuck线程数'}]})
		a = 1
		j = j + 6 							  
		#### get thread info ####
		for server in servers: 
			server_name = server.getName()
			threadpool_server = wl.getMBean('domainRuntime:/ServerRuntimes/' + server_name + '/ThreadPoolRuntime/ThreadPoolRuntime')
			if	threadpool_server:
				StandbyThreadCount = threadpool_server.getStandbyThreadCount()
				ExecuteThreadTotalCount = threadpool_server.getExecuteThreadTotalCount()
				ExecuteThreadIdleCount = threadpool_server.getExecuteThreadIdleCount()
				HoggingThreadCount = threadpool_server.getHoggingThreadCount()
				ExecuteThreads = threadpool_server.getExecuteThreads()
				activeThreadCount = ExecuteThreadTotalCount-StandbyThreadCount
				threadPercent = (activeThreadCount-ExecuteThreadIdleCount)/float(ExecuteThreadTotalCount)
				threadPercent = float("%.2f" % threadPercent)
				threadPercent = format(threadPercent,'.0%')
				stuckThreadCount = 0
				for stuckThread in ExecuteThreads:
					if(stuckThread.isStuck()):
						stuckThreadCount += 1
				data = [a,server_name,activeThreadCount,threadPercent,HoggingThreadCount,stuckThreadCount] 
				worksheet.write_row('A'+bytes(j), data)
			else:
				data = [a,server_name,'N/A','N/A','N/A','N/A']
				worksheet.write_row('A'+bytes(j), data)
				worksheet.write_comment('F'+bytes(j), ' server not running ')
			a += 1
			j += 1
		
		
		
		
		applications = wl.domainConfig().getAppDeployments() 
		worksheet.merge_range(xl_range(j+2, 0, j+3, 4), u'WebLogic应用程序配置及状态信息', domainformat)
		worksheet.add_table('A'+bytes(j+5)+':D'+bytes(len(applications) + j+6), {'columns': [{'header': u'序号'},{'header': u'应用程序名称'},{'header': u'应用程序target'},{'header': u'应用程序部署状态 '}]})
		b = 1
		j = j+6
		if	len(applications):
			for application in applications:
				for target in application.getTargets():#weblogic.management.configuration.AppDeploymentMBean
					targetname = target.getName()
					applicationinfo = wl.getMBean('domainConfig:/AppDeployments/' + application.getName())
					application_id =  applicationinfo.getApplicationIdentifier()
					application_runtime_mbean = wl.getMBean('domainRuntime:/AppRuntimeStateRuntime/AppRuntimeStateRuntime')
					application_status = application_runtime_mbean.getCurrentState(application_id, targetname)
					data = [b,application_id,targetname,application_status] #验证server_HealthState
					worksheet.write_row('A'+bytes(j), data)
					b += 1
					j += 1
		else:
			print " this is no application deploy! "
		


		
		jdbcsystemresources = wl.domainConfig().getJDBCSystemResources()
		worksheet.merge_range(xl_range(j+2, 0, j+3, 9), u'JDBC数据源连接数配置及使用信息', domainformat)		
		worksheet.add_table('A'+bytes(j+5)+':J'+bytes(len(jdbcsystemresources) + j+6+10), {'columns': [{'header': u'序号'},{'header': u'连接池名称'},{'header': u'连接池target'},{'header': u'连接池当前容量 '},{'header': u'连接池当前活动连接数 '},{'header': u'连接池当前活动连接最高数 '},{'header': u'当前连接池泄露书'},{'header': u'当前等待连接数 '},{'header': u'等待连接最高值'},{'header': u'连接池状态'}]})
		
		c = 1 
		j = j+6 
		if 	len(jdbcsystemresources):
			for jdbcsysresource in jdbcsystemresources:
				for target in jdbcsysresource.getTargets():	
					if target.getType() == 'Cluster':	
 						cluster = wl.getMBean('domainConfig:/Clusters/'+target.getName())
						for server in cluster.getServers():
							mbean_server = wl.getMBean('domainRuntime:/ServerRuntimes/'+server.getName())
							if 	mbean_server:
								jdbcdatasourceinfo = wl.getMBean('domainRuntime:/ServerRuntimes/'+server.getName()+'/JDBCServiceRuntime/'+server.getName()+'/JDBCDataSourceRuntimeMBeans/' + jdbcsysresource.getName())
								if jdbcdatasourceinfo:
									currcapacity = jdbcdatasourceinfo.getCurrCapacity()
									activeConnectionsCurrentCount = jdbcdatasourceinfo.getActiveConnectionsCurrentCount()
									activeConnectionsHighCount = jdbcdatasourceinfo.getActiveConnectionsHighCount()
									leakedconnectionCount = jdbcdatasourceinfo.getLeakedConnectionCount()
									waitingForConnectionCurrentCount = jdbcdatasourceinfo.getWaitingForConnectionCurrentCount()
									WaitingForConnectionHighCount = jdbcdatasourceinfo.getWaitingForConnectionHighCount()
									jdbcstate = jdbcdatasourceinfo.getState()
									data = [c,jdbcsysresource.getName(),target.getName(),currcapacity,activeConnectionsCurrentCount,activeConnectionsHighCount,leakedconnectionCount,		waitingForConnectionCurrentCount,WaitingForConnectionHighCount,jdbcstate]
									worksheet.write_row('A'+bytes(j), data)
								else:
									data = [c,jdbcsysresource.getName(),server.getName(),'N/A','N/A','N/A','N/A','N/A','N/A','N/A']
									worksheet.write_row('A'+bytes(j), data)
									worksheet.write_comment('J'+bytes(j), ' jdbc datasource not running')
							else:
								data = [c,jdbcsysresource.getName(),server.getName(),'N/A','N/A','N/A','N/A','N/A','N/A','N/A']
								worksheet.write_row('A'+bytes(j), data)
								worksheet.write_comment('J'+bytes(j), ' jdbc datasource not running')
							c += 1
							j += 1
					else:
						mbean_server = wl.getMBean('domainRuntime:/ServerRuntimes/'+target.getName())
					 
						
						if 	mbean_server:
							jdbcdatasourceinfo = wl.getMBean('domainRuntime:/ServerRuntimes/'+target.getName()+'/JDBCServiceRuntime/'+target.getName()+'/JDBCDataSourceRuntimeMBeans/' + jdbcsysresource.getName())
							if jdbcdatasourceinfo:
								currcapacity = jdbcdatasourceinfo.getCurrCapacity()
								activeConnectionsCurrentCount = jdbcdatasourceinfo.getActiveConnectionsCurrentCount()
								activeConnectionsHighCount = jdbcdatasourceinfo.getActiveConnectionsHighCount()
								leakedconnectionCount = jdbcdatasourceinfo.getLeakedConnectionCount()
								waitingForConnectionCurrentCount = jdbcdatasourceinfo.getWaitingForConnectionCurrentCount()
								WaitingForConnectionHighCount = jdbcdatasourceinfo.getWaitingForConnectionHighCount()
								jdbcstate = jdbcdatasourceinfo.getState()
								data = [c,jdbcsysresource.getName(),target.getName(),currcapacity,activeConnectionsCurrentCount,activeConnectionsHighCount,leakedconnectionCount,waitingForConnectionCurrentCount,WaitingForConnectionHighCount,jdbcstate]
								worksheet.write_row('A'+bytes(j), data)	
							else:
								data = [c,jdbcsysresource.getName(),target.getName(),'N/A','N/A','N/A','N/A','N/A','N/A','N/A']
								worksheet.write_row('A'+bytes(j), data)
								worksheet.write_comment('J'+bytes(j), ' jdbc datasource not running')
						else:
							data = [c,jdbcsysresource.getName(),target.getName(),'N/A','N/A','N/A','N/A','N/A','N/A','N/A']
							worksheet.write_row('A'+bytes(j), data)
							worksheet.write_comment('J'+bytes(j), ' jdbc datasource not running ')
							
						c += 1
						j += 1
					 
				
					#print 'jdbc .......2222222222222222 '
				
		#worksheet.merge_range(xl_range(j+2, 0, j+3, 4), u'文件句柄数', domainformat)		
		#worksheet.add_table('A'+bytes(j+5)+':D'+bytes(len(servers) + j+6), {'columns': [{'header': u'序号'},{'header': u'server名称'},{'header': u'打开文件句柄数'},{'header': u'最大文件句柄数'}]})
	except Exception, e:
		print '-> Exception occured!!!\n'

	worksheet.write('F6',time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())),ItemStyle)
	
	
# get jvm file descriptor
'''
def	getJVMFileDescriptorInfo(wlservername):
	current_fd_count = -1
	limit_fd_count = -1	
    wl.cd('domainCustom:/java.lang/java.lang:Location='+ wlservername +',type=OperatingSystem')
    os = string.lower(wl.cmo.get('Name'))
	print os
	if 'linux' in os or 'aix' in os or 'hp-ux' in os or 'sunos' in os or 'solaris' in os:
        current_fd_count = get('OpenFileDescriptorCount')
        limit_fd_count = get('MaxFileDescriptorCount')
		return current_fd_count, limit_fd_count
   jvm_current_fd_count, jvm_limit_fd_count = getJVMFileDescriptorInfo(wlserverName)
'''   

def writeXlsError(workbook,domainName):
	worksheet = workbook.add_worksheet(domainName)
	workformat = workbook.add_format()
	workformat.set_bold()
	workformat.set_font_color('red')
	workformat.set_font_size(20)
	workformat.set_align('center')
	workformat.set_align('vcenter')
	workformat.set_locked(True)
	
	workformat.set_border(2)
	worksheet.merge_range(xl_range(0, 0, 4, 10), u'WebLogic域('+domainName+u') 连接异常', workformat)	
	
	
	
if	__name__ == '__main__':
	
	with open('wls_domains_info.json','r')	as data_file:
		domains_config_info = json.load(data_file,encoding='utf-8',object_pairs_hook=OrderedDict)
		for business_name in domains_config_info:
			workbook = xlsxwriter.Workbook(checkfile+ "_"+business_name+'_'+time.strftime('%Y-%m-%d_%H_%M_%S',time.localtime(time.time()))+".xlsx")
			for domain_info in domains_config_info[business_name]:	
				try:
					writeXls(workbook,domain_info['domain_name'],domain_info['admin_username'],domain_info['admin_password'],domain_info['admin_url'])
				except:
					print '-> connect error !!!\n'
					writeXlsError(workbook,domain_info['domain_name'])
					continue
			workbook.close()

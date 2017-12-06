<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>
    <base href="<%=basePath%>">
    <title>POI</title>
	<meta http-equiv="pragma" content="no-cache">
	<meta http-equiv="cache-control" content="no-cache">
	<meta http-equiv="expires" content="0">    
	<meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
	<meta http-equiv="description" content="This is my page">
	<!--
	<link rel="stylesheet" type="text/css" href="styles.css">
	-->
<script type="text/javascript">
function formSubmit(mapping){
	var postForm = document.getElementById("excelForm");
	postForm.action = "<%=basePath%>" + mapping;
	postForm.submit();
}
</script>
  </head>
  <body>
  	<h2>文件导出</h2><hr><br>
  	<form id="excelForm" method="post">
  		日&nbsp;&nbsp;&nbsp;&nbsp;期&nbsp;&nbsp;&nbsp;: <input name="queryDate" value = ""/><br>
  		文件名称: <input name="fileName" value = ""/><br>
	    <a href="javascript:void(0)" onclick="formSubmit('replace')">导出数据</a><br><br>
    </form>
    
    
  	<h2>文件上传</h2><hr><br>
    <form method="post" action="/poi/uploadServlet" enctype="multipart/form-data">
	    选择一个文件:
	    <input type="file" name="uploadFile" />
	    <br/><br/>
	    <input type="submit" value="上传" />
	</form>
  </body>
</html>

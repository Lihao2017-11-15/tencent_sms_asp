<!--#include file="comm_sms.asp"-->
<%
'------------------------------------------------
' Copyright (c) 2020, shenzhen haowen
'
' SPDX-License-Identifier: Apache-2.0
'
' Change Logs:
' Date           Author       		Notes
' 2020-12-05     812256@qq.com    	first version
'------------------------------------------------
'重要说明：
'1.在腾讯平台注册审核
'2.提交短信模板审核
'3.通过后再根据短信模板实际参数测试
'
'-------需要设置的部分开始----------
strMobile = "查看腾讯短信平台接口"
strAppKey = "查看腾讯短信平台接口"
sign =      "查看腾讯短信平台接口"
tpl_id = 561230 '查看腾讯短信平台接口
'-------需要设置的部分结束----------

RANDOMIZE
strRand = cstr(int(rnd()*9999999999))
strTime = cstr(ToUnixTime(now()))
sig = sha256("appkey=" & strAppKey & "&random=" & strRand & "&time=" & strTime & "&mobile=" & strMobile)

phone =  Request.Form("p")
cartype = Request.Form("c")
otime = Request.Form("o")
addr = Request.Form("a")
if phone=0 or cartype="" or otime="" or addr="" then
%>
<html>
<head>
	<style>
		label {
			display: block;
			padding: 0.5em;
			max-width: 20em;
		}
		input {
			float: right;
		}
	</style>
	<title>腾讯短信平台ASP接口测试</title>
</head>
<body>
	<form method="post" action="">
		<p><b>腾讯短信平台ASP接口测试</b></p>
		<label>接收手机号：<input name="p" value="<%=strMobile%>"></label>
		<label>短信模板参数1：<input name="c" value="cartype1"></label>
		<label>短信模板参数2：<input name="o" value="2020-12-05"></label>
		<label>短信模板参数3：<input name="a" value="预约地点"></label>
		<label>短信模板参数4：<input type="submit" value="提交测试"></label>
	</form>
</body>
</html>
<%
else
	%>
	<!-- 参考：https://cloud.tencent.com/document/product/382/5976 -->
	<script src="vendor/jquery/jquery.min.js"></script>
	<script>
		$(function () {

			postdata = {
				"ext": "",
				"extend": "",
				"params": [
					//这里是你短信模板中的实际参数
					"<%=phone%>",
					"<%=cartype%>",
					"<%=otime%>",
					"<%=addr%>"
				],
				"sig": "<%=sig%>",
				"sign": "<%=sign%>",
				"tel": {
					"mobile": "<%=phone%>",
					"nationcode": "86"
				},
				"time": "<%=strTime%>",
				"tpl_id": <%=tpl_id%>
			};
			console.log(postdata);
			console.log(JSON.stringify(postdata));
			$.ajax({
				contentType: 'application/json',
				type: 'POST',
				url: "https://yun.tim.qq.com/v5/tlssmssvr/sendsms?sdkappid=1400338975&random=<%=strRand%>",
				dataType: "json",
				data: JSON.stringify(postdata),
				success: function (message) {
					if (message.result != 0) {
						alert("短信发送失败：" + message.errmsg + "，错误码：" + message.result + "\n\n看到此信息说明接口调用成功，但是参数设置错误");
						history.back(-1);
					}
					else {
						alert("通知短信发送成功。");
					}
				},
				error: function (message) {
					alert("短信发送失败：", message);
				}
			});
		});
	</script>
<%end if%>
<%
Const s_lang_meta = "<META HTTP-EQUIV=""Content-Type"" content=""text/html; charset=gb_2312-80"">"
Const s_lang_return = "返回"
Const s_lang_exit = "退出"
Const s_lang_modify = "修改"
Const s_lang_close = "关闭"
Const s_lang_minute = "分钟"
Const s_lang_hour = "小时"
Const s_lang_day = "天"
Const s_lang_save = "保存"
Const s_lang_add = "添加"
Const s_lang_del = "删除"
Const s_lang_tpf = "支持通配符方式. (<font color=""#901111"">*</font>: 任意长度的任何内容.&nbsp;&nbsp;<font color=""#901111"">?</font>: 一个字符的任何内容.)"
Const s_lang_not_recom = "副作用较大, 不建议使用"
Const s_lang_recom = "建议选中"
Const s_lang_setting = "设置"
Const s_lang_inputerr = "输入错误!"
Const s_lang_enable = "启动项设置"
Const s_lang_msg = "内容"
Const s_lang_find = "查找"
Const s_lang_all = "所有"
Const s_lang_cancel = "取消"
Const s_lang_font = "宋体,MS SONG,SimSun,tahoma,sans-serif"

Const s_lang_0001 = "确实要清空动态信任IP列表吗?"
Const s_lang_0002 = "确实要清空动态不良IP列表吗?"
Const s_lang_0003 = "启用灰名单防垃圾邮件功能"
Const s_lang_0004 = "清空动态信任IP列表"
Const s_lang_0005 = "灰名单功能"
Const s_lang_0006 = "灰名单重试拒绝时间"
Const s_lang_0007 = "灰名单记录保存期"
Const s_lang_0008 = "动态信任IP列表保存期"
Const s_lang_0009 = "动态不良IP列表保存期"
Const s_lang_0010 = "清空动态不良IP列表"
Const s_lang_0011 = "此功能可以有效识别和控制来自邮件群发软件发送到系统内的垃圾邮件."
Const s_lang_0012 = "使用缺省设置"
Const s_lang_0013 = "设置可信任IP"
Const s_lang_0014 = "设置可信任发件人"
Const s_lang_0015 = "设置可直接通过灰名单验证的IP地址."
Const s_lang_0016 = "设置灰名单可信任IP"
Const s_lang_0017 = "设置灰名单可信任发件人"
Const s_lang_0018 = "设置可直接通过灰名单验证的发件人."

Const s_lang_0021 = "设置灰名单防垃圾邮件功能"
Const s_lang_0022 = ""

Const s_lang_0023 = "设置信任邮件地址"
Const s_lang_0024 = "确保来自信任邮件地址的邮件可进入收件箱中"
Const s_lang_0025 = "启用信任邮件地址功能"

Const s_lang_0026 = "设置SMTP免认证IP"
Const s_lang_0027 = "无需SMTP认证即可发信的IP地址集"
Const s_lang_0028 = "启用SMTP免认证IP功能"

Const s_lang_0029 = "启用邮件防伪造功能"
Const s_lang_0030 = "检查邮件头 From 邮件地址中域名的 A、MX、SPF 记录"

Const s_lang_0031 = "排首位的域名为主域, 主域内的用户即使为含域名的帐号, 登录时也可以只输入@号前的部分做为帐号"

Const s_lang_0032 = "设置LDAP资料"
Const s_lang_0033 = "修改LDAP密码"
Const s_lang_0034 = "帐户名"
Const s_lang_0035 = "姓名"
Const s_lang_0036 = "电子邮件地址"
Const s_lang_0037 = "家庭电话"
Const s_lang_0038 = "寻呼机"
Const s_lang_0039 = "移动电话"
Const s_lang_0040 = "商务电话"
Const s_lang_0041 = "商务传真"
Const s_lang_0042 = "职务"
Const s_lang_0043 = "部门"
Const s_lang_0044 = "办公室"
Const s_lang_0045 = "姓"
Const s_lang_0046 = "名"
Const s_lang_0047 = "输入的密码不相同"
Const s_lang_0048 = "测试帐号不允许修改密码"
Const s_lang_0049 = "新 密 码"
Const s_lang_0050 = "密码确认"
Const s_lang_0051 = "修改邮箱密码时, 同时更改LDAP的登录密码"
Const s_lang_0052 = "允许加入到企业地址簿"
Const s_lang_0053 = "允许用户修改数据"
Const s_lang_0054 = "允许用户数据加入到LDAP服务"
Const s_lang_0055 = "允许用户登录LDAP服务"
Const s_lang_0056 = "允许用户修改所属部门"
Const s_lang_0057 = "收缩"
Const s_lang_0058 = "展开"
Const s_lang_0059 = "确认"
Const s_lang_0060 = "的LDAP设置"
Const s_lang_0061 = "设置LDAP密码"
Const s_lang_0062 = "(为空时不修改原密码)"
Const s_lang_0063 = "LDAP设置"
Const s_lang_0064 = "管理扩展LDAP数据"
Const s_lang_0065 = "[新数据]"
Const s_lang_0066 = "UID中包含非法字符."
Const s_lang_0067 = "文件名"
Const s_lang_0068 = "已存在相同UID数据"
Const s_lang_0069 = "-注释对应项-"
Const s_lang_0070 = "LDAP数据"
Const s_lang_0071 = "启用LDAP功能"
Const s_lang_0072 = "组织"

Const s_lang_0073 = "允许在SMTP命令中应答无此用户时的情况"
Const s_lang_0074 = "允许接收外部邮件"
Const s_lang_0075 = "允许自动回复功能"
Const s_lang_0076 = "允许接收外部发来的邮件"
Const s_lang_0077 = "禁止接收外部发来的邮件"

Const s_lang_0078 = "邮件群发设置"
Const s_lang_0079 = "设置邮件群发功能"
Const s_lang_0080 = "显示为邮件群发地址发送"
Const s_lang_0081 = "开启授权使用邮件群发的功能"
Const s_lang_0082 = "键入授权邮址(支持通配符)"
Const s_lang_0083 = "授权邮址列表"
Const s_lang_0084 = "开启授权功能后, 只有从授权邮件地址发到群发帐号的邮件才会被系统群发"

Const s_lang_0085 = "创建新邮件列表"
Const s_lang_0086 = "刷新"
Const s_lang_0087 = "邮件列表"
Const s_lang_0088 = "序号"
Const s_lang_0089 = "邮件列表发信人"
Const s_lang_0090 = "显示列<br>表邮址"
Const s_lang_0091 = "私密"
Const s_lang_0092 = "创建者邮件地址"
Const s_lang_0093 = "编辑"
Const s_lang_0094 = "[管理员]"
Const s_lang_0095 = "邮件列表接收用户不可为空."
Const s_lang_0096 = "邮件列表发送人不可为空."
Const s_lang_0097 = "只能添加系统内用户"
Const s_lang_0098 = "编辑邮件列表"
Const s_lang_0099 = "邮件列表发送人"
Const s_lang_0100 = "显示为此邮件列表地址发送"
Const s_lang_0101 = "私密邮件列表 (非此列表成员不能使用)"
Const s_lang_0102 = "所有用户"
Const s_lang_0103 = "接收邮件列表的用户"
Const s_lang_0104 = "登录IP限制 (需要设置帐号保护功能)"
Const s_lang_0105 = "开启登录IP限制功能"
Const s_lang_0106 = "限制登录IP地址"
Const s_lang_0107 = "信任超过"
Const s_lang_0108 = "的邮件不是垃圾信"

Const s_lang_0109 = "垃圾邮件投诉审核"
Const s_lang_0110 = "审核用户举报的垃圾邮件信息"
Const s_lang_0111 = "垃圾邮件数据库管理"
Const s_lang_0112 = "可以将误判的邮件从垃圾邮件数据库中移除"
Const s_lang_0113 = "<b>选中的不是垃圾邮件</b>(其他都是垃圾邮件)"
Const s_lang_0114 = "<b>选中的是垃圾邮件</b>(其他都不是垃圾邮件)"
Const s_lang_0115 = "确实要删除吗?"

Const s_lang_0116 = "备份选项"
Const s_lang_0117 = "不备份"
Const s_lang_0118 = "备份Web发送邮件"
Const s_lang_0119 = "备份所有发送邮件"
Const s_lang_0120 = "邮箱即时清理选项"
Const s_lang_0121 = "邮箱满时允许清空垃圾箱"
Const s_lang_0122 = "邮箱满时允许清理发件箱"

Const s_lang_0123 = "邮件撤回"
Const s_lang_0124 = "有"
Const s_lang_0125 = "封可撤回邮件"
Const s_lang_0126 = "状态"
Const s_lang_0127 = "主题"
Const s_lang_0128 = "日期"
Const s_lang_0129 = "[无主题]"
Const s_lang_0130 = "紧急邮件"
Const s_lang_0131 = "慢件"
Const s_lang_0132 = "已发送完毕"
Const s_lang_0133 = "正在发送"
Const s_lang_0134 = "日期 : 今天"
Const s_lang_0135 = "日期 : 昨天"
Const s_lang_0136 = "日期 : 前天"
Const s_lang_0137 = "日期 : 三天前"
Const s_lang_0138 = "日期 : 一周前"
Const s_lang_0139 = "年"
Const s_lang_0140 = "月"
Const s_lang_0141 = "日"
Const s_lang_0142 = "确实要撤回邮件吗?"
Const s_lang_0143 = "撤回已选"
Const s_lang_0144 = "撤回全部"
Const s_lang_0145 = "优先级"
Const s_lang_0146 = "普通邮件"
Const s_lang_0147 = "发件人"
Const s_lang_0148 = "发件人地址"
Const s_lang_0149 = "收件人"
Const s_lang_0150 = "撤回"
Const s_lang_0151 = "群发邮件"
Const s_lang_0152 = "返回顶部"
Const s_lang_0153 = "正在进行撤回处理"
Const s_lang_0154 = "撤回成功"
Const s_lang_0155 = "撤回失败"
Const s_lang_0156 = "撤回过程中出现故障"
Const s_lang_0157 = "对方不支持撤回"
Const s_lang_0158 = "正常"
Const s_lang_0159 = "启用邮件撤回功能"
Const s_lang_0160 = "待撤回邮件的保留天数"
Const s_lang_0161 = "1. 支持对发往本系统以及发往其他企业版本WinWebMail Server邮局邮件的撤回."
Const s_lang_0162 = "2. 如果撤回成功, 对方收到的邮件将被清除, 并且对方会得到邮件已被撤回的提示."
Const s_lang_0163 = "3. 如果对方已经阅读, 将无法撤回."
Const s_lang_0164 = "显示邮件接收地址信息"
Const s_lang_0165 = "邮件撤回通知"
Const s_lang_0166 = "主　题："
Const s_lang_0167 = "宏变量"
Const s_lang_0168 = "表示来信的日期"
Const s_lang_0169 = "表示来信的时间"
Const s_lang_0170 = "表示来信的发件人名称"
Const s_lang_0171 = "表示来信的发件人邮件地址"
Const s_lang_0172 = "表示来信的标题"

Const s_lang_0173 = "禁用此邮件列表"
Const s_lang_0174 = "此列表<br>被禁用"

Const s_lang_0175 = "监控到"
Const s_lang_0176 = "的已接收邮件共"
Const s_lang_0177 = "的已发送邮件共"
Const s_lang_0178 = "封"
Const s_lang_0179 = "长度"
Const s_lang_0180 = "发送数"
Const s_lang_0181 = "系统邮件"
Const s_lang_0182 = "邮件"
Const s_lang_0183 = "数字签名邮件"
Const s_lang_0184 = "数字加密邮件"
Const s_lang_0185 = "未开启对"
Const s_lang_0186 = "用户的邮件发送监控, 现在就"
Const s_lang_0187 = "设置"
Const s_lang_0188 = "未开启系统级邮件外发超限额自动监控功能, 现在就"
Const s_lang_0189 = "开启"
Const s_lang_0190 = "字节"

Const s_lang_0191 = "密码输入错误"
Const s_lang_0192 = "您输入的密码强度不足"
Const s_lang_0193 = "新密码不可与原密码相同"
Const s_lang_0194 = "修改密码"
Const s_lang_0195 = "太短"
Const s_lang_0196 = "弱"
Const s_lang_0197 = "一般"
Const s_lang_0198 = "极佳"
Const s_lang_0199 = "尊敬的用户，为了确保您的帐号安全，请立即修改登录密码"
Const s_lang_0200 = "密码强度"
Const s_lang_0201 = "退出邮箱"
Const s_lang_0202 = "正在验证, 请稍候"

Const s_lang_0203 = "邮件外发排行"
Const s_lang_0204 = "为防止服务器IP地址被列入黑名单, 需要对用户外发邮件进行控制"

Const s_lang_0205 = "启用外发邮件排行统计功能"
Const s_lang_0206 = "排行统计数据的保留天数"
Const s_lang_0207 = "启用邮件外发超限额自动监控功能"
Const s_lang_0208 = "开启自动监控功能时的每日最大外发邮件数量"
Const s_lang_0209 = "自动监控功能持续天数"

Const s_lang_0210 = "开始查找"
Const s_lang_0211 = "开启外发监控功能"
Const s_lang_0212 = "暂无数据"
Const s_lang_0213 = "日期列表"
Const s_lang_0214 = "请选中待检查的日期"
Const s_lang_0215 = "未找到指定内容"
Const s_lang_0216 = "配置模板"
Const s_lang_0217 = "应用模板到所选用户"
Const s_lang_0218 = "页内查找"
Const s_lang_0219 = "排名"
Const s_lang_0220 = "用户帐号"
Const s_lang_0221 = "外发数量"
Const s_lang_0222 = "查看监控"
Const s_lang_0223 = "高级设置"
Const s_lang_0224 = "发送"
Const s_lang_0225 = "接收"
Const s_lang_0226 = "配置"

Const s_lang_0227 = "配置系统模板"
Const s_lang_0228 = "恢复默认设置"
Const s_lang_0229 = "邮件发送监控"
Const s_lang_0230 = "此模板生效"
Const s_lang_0231 = "对用户发送的邮件进行监控"
Const s_lang_0232 = "仅监控对系统外发送的邮件"
Const s_lang_0233 = "的邮件"
Const s_lang_0234 = "仅监控小于"
Const s_lang_0235 = "被监控邮件的保存天数"
Const s_lang_0236 = "结束日期为"
Const s_lang_0237 = "永久有效"
Const s_lang_0238 = "限制用户每日外发邮件数量"
Const s_lang_0239 = "每日外发邮件数上限"
Const s_lang_0240 = "强制用户修改密码"
Const s_lang_0241 = "取消强制修改密码"
Const s_lang_0242 = "要求密码强度: 低"
Const s_lang_0243 = "要求密码强度: 中"
Const s_lang_0244 = "要求密码强度: 高"
Const s_lang_0245 = "禁用自动回复"
Const s_lang_0246 = "禁用自动转发"
Const s_lang_0247 = "禁止用户对系统外发送邮件"
Const s_lang_0248 = "禁用用户帐号"
Const s_lang_0249 = "邮件接收监控"
Const s_lang_0250 = "对用户接收的邮件进行监控"
Const s_lang_0251 = "不限大小"
Const s_lang_0252 = "只有选中""此模板生效""后, 该项模板的设置才会被应用到指定用户."

Const s_lang_0253 = "等于"
Const s_lang_0254 = "不等于"
Const s_lang_0255 = "包含"
Const s_lang_0256 = "不包含"
Const s_lang_0257 = "通配符等于"
Const s_lang_0258 = "等于"
Const s_lang_0259 = "大于"
Const s_lang_0260 = "小于"
Const s_lang_0261 = "通配符等于"
Const s_lang_0262 = "用户"
Const s_lang_0263 = "应用系统模板"
Const s_lang_0264 = "邮件发送过滤"
Const s_lang_0265 = "对用户发送邮件进行过滤"
Const s_lang_0266 = "仅过滤小于"

Const s_lang_0267 = "今天"
Const s_lang_0268 = "凌晨"
Const s_lang_0269 = "上午"
Const s_lang_0270 = "下午"
Const s_lang_0271 = "晚上"
Const s_lang_0272 = "条发信记录"
Const s_lang_0273 = "投递完成"
Const s_lang_0274 = "投递中"
Const s_lang_0275 = "查看"
Const s_lang_0276 = "查看邮件发送状态"
Const s_lang_0277 = "投递成功"
Const s_lang_0278 = "投递失败"
Const s_lang_0279 = "发信查询"
Const s_lang_0280 = "发送状态"
Const s_lang_0281 = "用户登录WebMail, IP:"
Const s_lang_0282 = "用户修改密码, IP:"
Const s_lang_0283 = "中午"

Const s_lang_0284 = "从中转站添加"
Const s_lang_0285 = "请稍候"
Const s_lang_0286 = "共有"
Const s_lang_0287 = "个附件"
Const s_lang_0288 = "上传附件"
Const s_lang_0289 = "错误"
Const s_lang_0290 = "查看中转站"
Const s_lang_0291 = "查看全部"
Const s_lang_0292 = "中转站"

Const s_lang_0293 = "所有接收者均通过验证"
Const s_lang_0294 = "以下接收者未通过验证"

Const s_lang_0295 = "搜索文件名"
Const s_lang_mh = "："
Const s_lang_jh = "。"
Const s_lang_charset = "gb_2312-80"
Const s_lang_html = ""
Const s_lang_0296 = "网络存储"
Const s_lang_0297 = "启用链接式附件功能"
Const s_lang_0298 = "链接式附件下载地址"
Const s_lang_0299 = "附件有效天数"

Const s_lang_0300 = "标签管理"
Const s_lang_0301 = "改名"
Const s_lang_0302 = "新建标签"
Const s_lang_0303 = "请您输入标签名称"
Const s_lang_0304 = "标签名称不能为空"
Const s_lang_0305 = "存在同名的标签"
Const s_lang_0306 = "重命名标签"
Const s_lang_0307 = "请您输入新的标签名称"
Const s_lang_0308 = "新建标签"
Const s_lang_0309 = "标签文件夹"
Const s_lang_0310 = "未读邮件"
Const s_lang_0311 = "总邮件"
Const s_lang_0312 = "操作"
Const s_lang_0313 = "确定"

Const s_lang_0314 = "您的操作<b>成功</b>"
Const s_lang_0315 = "您的操作<b>失败</b>"

Const s_lang_0316 = "启用登录时验证码功能"

Const s_lang_0317 = "通讯组"
Const s_lang_0318 = "搜索"
Const s_lang_0319 = "通讯组名称"
Const s_lang_0320 = "邮件地址"
Const s_lang_0321 = "地址簿"
Const s_lang_0322 = "系统公共地址簿"
Const s_lang_0323 = "域公共地址簿"
Const s_lang_0324 = "创建新地址"
Const s_lang_0325 = "删除地址"

Const s_lang_0326 = "写邮件"
Const s_lang_0327 = "收件箱"
Const s_lang_0328 = "邮件查找"
Const s_lang_0329 = "星标邮件"
Const s_lang_0330 = "企业地址簿"
Const s_lang_0331 = "我的文件夹"
Const s_lang_0332 = "草稿箱"
Const s_lang_0333 = "已发送"
Const s_lang_0334 = "垃圾箱"
Const s_lang_0335 = "标签"
Const s_lang_0336 = "我的存储文件夹"
Const s_lang_0337 = "其他功能"
Const s_lang_0338 = "数字证书"
Const s_lang_0339 = "效率手册"
Const s_lang_0340 = "共享文件夹"
Const s_lang_0341 = "投票"
Const s_lang_0342 = "公共文件夹"
Const s_lang_0343 = "选项"
Const s_lang_0344 = "记事本"
Const s_lang_0345 = "系统设置"
Const s_lang_0346 = "用户管理"
Const s_lang_0347 = "域设置"
Const s_lang_0348 = "域用户管理"
Const s_lang_0349 = "发信查询"
Const s_lang_0350 = "地址簿"

Const s_lang_0351 = "请输入“通讯组名称”"
Const s_lang_0352 = "请输入“邮件地址”"
Const s_lang_0353 = "成员"
Const s_lang_0354 = "若要将成员添加到通讯组，请键入他们的电子邮件地址（每个地址之间用逗号分隔）。"
Const s_lang_0355 = "从私人地址簿添加"

Const s_lang_0356 = "文件类型输入错误"
Const s_lang_0357 = "确实要拷贝到 [系统公共地址簿] 吗"
Const s_lang_0358 = "确实要拷贝到 [域公共地址簿] 吗"
Const s_lang_0359 = "上传文件"
Const s_lang_0360 = "昵称"
Const s_lang_0361 = "查看"
Const s_lang_0362 = "查看数字证书"
Const s_lang_0363 = "创建新地址"
Const s_lang_0364 = "删除地址"
Const s_lang_0365 = "写信"
Const s_lang_0366 = "导入 csv 文件"
Const s_lang_0367 = "导出私人地址簿"
Const s_lang_0368 = "拷贝到[域公共地址簿]"
Const s_lang_0369 = "拷贝到[系统公共地址簿]"
Const s_lang_0370 = "确实要拷贝到 [私人地址簿] 吗"
Const s_lang_0371 = "数字证书"
Const s_lang_0372 = "收藏"
Const s_lang_0373 = "下载"
Const s_lang_0374 = "拷贝到[私人地址簿]"
Const s_lang_0375 = "保存并继续添加"
Const s_lang_0376 = "请输入“昵称”"
Const s_lang_0377 = "Email地址"
Const s_lang_0378 = "姓"
Const s_lang_0379 = "名"
Const s_lang_0380 = "其他邮件地址"
Const s_lang_0381 = "公司"
Const s_lang_0382 = "通讯"
Const s_lang_0383 = "家庭电话"
Const s_lang_0384 = "工作电话"
Const s_lang_0385 = "移动电话"
Const s_lang_0386 = "地址信息"
Const s_lang_0387 = "邮政编码"
Const s_lang_0388 = "地址"
Const s_lang_0389 = "城市"
Const s_lang_0390 = "省"
Const s_lang_0391 = "国家"
Const s_lang_0392 = "其他信息"
Const s_lang_0393 = "生日"
Const s_lang_0394 = "主页"

Const s_lang_0395 = "退出"
Const s_lang_0396 = "失败"
Const s_lang_0397 = "密码错误"
Const s_lang_0398 = "文件夹不存在或不允许访问"
Const s_lang_0399 = "确实要彻底删除吗?"
Const s_lang_0400 = "确实要清空当前文件夹中的所有文件吗?"
Const s_lang_0401 = "(共"
Const s_lang_0402 = " 个文件)"
Const s_lang_0403 = "全部清空"
Const s_lang_0404 = "移动到"
Const s_lang_0405 = "阅读"
Const s_lang_0406 = "文件名"
Const s_lang_0407 = "说明"
Const s_lang_0408 = "日期"
Const s_lang_0409 = "长度"
Const s_lang_0410 = "[无主题]"
Const s_lang_0411 = "未读"
Const s_lang_0412 = "已读"
Const s_lang_0413 = "保存到网络存储中"
Const s_lang_0414 = "正在上传, 请稍候..."
Const s_lang_0415 = "全选"
Const s_lang_0416 = "不选"
Const s_lang_0417 = "反选"
Const s_lang_0418 = "未读"
Const s_lang_0419 = "已读"
Const s_lang_0420 = "邮件"
Const s_lang_0421 = "新邮件"
Const s_lang_0422 = "要将所选邮件报告为垃圾邮件吗?"
Const s_lang_0423 = " 封邮件)"
Const s_lang_0424 = "彻底删除"
Const s_lang_0425 = "标记为"
Const s_lang_0426 = "更多"
Const s_lang_0427 = "已回复邮件"
Const s_lang_0428 = "已转发邮件"
Const s_lang_0429 = "新系统邮件"
Const s_lang_0430 = "发件箱"
Const s_lang_0431 = "已读邮件"
Const s_lang_0432 = "未读邮件"
Const s_lang_0433 = "星标邮件"
Const s_lang_0434 = "取消星标"
Const s_lang_0435 = "举报为垃圾邮件"
Const s_lang_0436 = "收缩"
Const s_lang_0437 = "展开"
Const s_lang_0438 = "日期 : 一个月前"
Const s_lang_0439 = "日期 : 一年前"
Const s_lang_0440 = "确实要清空垃圾箱吗?"
Const s_lang_0441 = "收件人"
Const s_lang_0442 = "显示为：会话模式"
Const s_lang_0443 = "确实要彻底删除会话吗?"
Const s_lang_0444 = "显示为：普通模式"
Const s_lang_0445 = "查找到："
Const s_lang_0446 = " 封邮件"
Const s_lang_0447 = "第"
Const s_lang_0448 = "行输入错误"
Const s_lang_0449 = "个人资料"
Const s_lang_0450 = "帐号到期日期"
Const s_lang_0451 = "注册信息"
Const s_lang_0452 = "要将此邮件报告为垃圾邮件吗?"
Const s_lang_0453 = "操作成功."
Const s_lang_0454 = "操作失败."
Const s_lang_0455 = "上一封"
Const s_lang_0456 = "下一封"
Const s_lang_0457 = "隐藏"
Const s_lang_0458 = "邮件详情"
Const s_lang_0459 = "设置星标"
Const s_lang_0460 = "回复"
Const s_lang_0461 = "回复全部"
Const s_lang_0462 = "转发"
Const s_lang_0463 = "再次发送"
Const s_lang_0464 = "[数字签名邮件] 未知的数字证书"
Const s_lang_0465 = "[数字加密邮件] 未知的数字证书"
Const s_lang_0466 = "[数字签名邮件] 有效的数字证书"
Const s_lang_0467 = "[数字加密邮件] 有效的数字证书"
Const s_lang_0468 = "[数字签名邮件] 无效的数字证书"
Const s_lang_0469 = "[数字加密邮件] 无效的数字证书"
Const s_lang_0470 = "数字证书"
Const s_lang_0471 = "数字签名"
Const s_lang_0472 = "未知的数字证书"
Const s_lang_0473 = "签名时间"
Const s_lang_0474 = "邮件已被加密</font>] 请"
Const s_lang_0475 = "上传您的私人数字证书"
Const s_lang_0476 = "邮件已被加密"
Const s_lang_0477 = "请输入您数字证书的密码"
Const s_lang_0478 = "您的数字证书无法解密"
Const s_lang_0479 = "发件人"
Const s_lang_0480 = "加入地址簿"
Const s_lang_0481 = "拒收"
Const s_lang_0482 = "日&nbsp;&nbsp;期"
Const s_lang_0483 = "查看"
Const s_lang_0484 = "大字"
Const s_lang_0485 = "中字"
Const s_lang_0486 = "小字"
Const s_lang_0487 = "大&nbsp;&nbsp;小"
Const s_lang_0488 = "年"
Const s_lang_0489 = "月"
Const s_lang_0490 = "日"
Const s_lang_0491 = "点"
Const s_lang_0492 = ""
Const s_lang_0493 = "定时发送"
Const s_lang_0494 = "更改"
Const s_lang_0495 = "超文本邮件"
Const s_lang_0496 = "浏览超文本格式邮件"
Const s_lang_0497 = "收件人"
Const s_lang_0498 = "投&nbsp;&nbsp;诉"
Const s_lang_0499 = "报告为垃圾邮件"
Const s_lang_0500 = "这不是垃圾邮件"
Const s_lang_0501 = "优先级"
Const s_lang_0502 = "抄送地址"
Const s_lang_0503 = "邮件头"
Const s_lang_0504 = "显示邮件头"
Const s_lang_0505 = "附件"
Const s_lang_0506 = "转存到网络存储"
Const s_lang_0507 = "类型"
Const s_lang_0508 = "活动邀请函"
Const s_lang_0509 = "活动提醒函"
Const s_lang_0510 = "生日"
Const s_lang_0511 = "提醒"
Const s_lang_0512 = "重复"
Const s_lang_0513 = "by + ""年"" + bm + ""月"" + bd + ""日 "" + convWeeekName(currentDate.getDay())"
Const s_lang_0514 = "星期日"
Const s_lang_0515 = "星期一"
Const s_lang_0516 = "星期二"
Const s_lang_0517 = "星期三"
Const s_lang_0518 = "星期四"
Const s_lang_0519 = "星期五"
Const s_lang_0520 = "星期六"
Const s_lang_0521 = "全天"
Const s_lang_0522 = "活动将于"
Const s_lang_0523 = "小时"
Const s_lang_0524 = "分钟后开始"
Const s_lang_0525 = "活动已经开始"
Const s_lang_0526 = "活动已经结束"
Const s_lang_0527 = "此活动已设定重复功能"
Const s_lang_0528 = "查看并回复此邀请函"" style=""WIDTH: 150px"
Const s_lang_0529 = "查看我的活动邀请列表"" style=""WIDTH: 160px"
Const s_lang_0530 = "查看活动当天的安排"" style=""WIDTH: 150px"
Const s_lang_0531 = "提醒"
Const s_lang_0532 = "查看此活动详细内容"" style=""WIDTH: 150px"
Const s_lang_0533 = "事件信息"
Const s_lang_0534 = "活动发起人"
Const s_lang_0535 = "发起人邮址"
Const s_lang_0536 = "日期"
Const s_lang_0537 = "时间"
Const s_lang_0538 = "重复"
Const s_lang_0539 = "此活动有设定重复功能"
Const s_lang_0540 = "位置"
Const s_lang_0541 = "城市"
Const s_lang_0542 = "地址"
Const s_lang_0543 = "电话"
Const s_lang_0544 = "其他"
Const s_lang_0545 = "您的请柬发送于"
Const s_lang_0546 = "转递"
Const s_lang_0547 = "单独查看"
Const s_lang_0548 = "查看附件"
Const s_lang_0549 = ""
Const s_lang_0550 = "星标邮件"
Const s_lang_0551 = ""
Const s_lang_0552 = ""
Const s_lang_0553 = "gb_2312-80"
Const s_lang_0554 = "隐藏收件地址"
Const s_lang_0555 = "展示收件地址"
Const s_lang_0556 = "发送次数"
Const s_lang_0557 = "收件人地址"
Const s_lang_0558 = "<b>失败</b>"
Const s_lang_0559 = "您的权限不足"

Const s_lang_0560 = "RBL黑名单检查"
Const s_lang_0561 = "可以实时检查服务器IP地址是否被列入黑名单并通知管理员"
Const s_lang_0562 = "启用RBL黑名单检查功能"
Const s_lang_0563 = "键入通知邮件地址 (系统内)"
Const s_lang_0564 = "通知邮件地址列表"
Const s_lang_0565 = "键入待检查IP"
Const s_lang_0566 = "待检查IP列表"
Const s_lang_0567 = "键入RBL查询网站"
Const s_lang_0568 = "RBL查询网站列表"
Const s_lang_0569 = "常用的RBL查询网站有：<br>sbl.spamhaus.org<br>cbl.abuseat.org<br>bl.spamcop.net<br>pbl.spamhaus.org<br>dnsbl.sorbs.net<br>"

Const s_lang_0570 = "文件归档"
Const s_lang_0571 = "归档文件夹"
Const s_lang_0572 = "要将所选邮件移到归档文件夹吗?"
Const s_lang_0573 = "主题"
Const s_lang_0574 = "发信地址"
Const s_lang_0575 = "发信人"
Const s_lang_0576 = "包含"
Const s_lang_0577 = "等于"
Const s_lang_0578 = "等于(支持通配符)"
Const s_lang_0579 = "要将邮件移到归档文件夹吗?"
Const s_lang_0580 = "是否删除"
Const s_lang_0581 = "年"
Const s_lang_0582 = "月数据"
Const s_lang_0583 = "文档总数："
Const s_lang_0584 = "已关闭"
Const s_lang_0585 = "限额："
Const s_lang_0586 = ""
Const s_lang_0587 = "年归档数据"

Const s_lang_0588 = "邮件审核"
Const s_lang_0589 = "同意发送"
Const s_lang_0590 = "拒绝发送"
Const s_lang_0591 = "同意所选"
Const s_lang_0592 = "拒绝所选"
Const s_lang_0593 = "全部同意"
Const s_lang_0594 = "全部拒绝"
Const s_lang_0595 = "同意发送当前所有待审核邮件吗?"
Const s_lang_0596 = "拒绝发送当前所有待审核邮件吗?"
Const s_lang_0597 = "可以指定审核管理员对用户的外发邮件进行审批"
Const s_lang_0598 = "启用邮件审核功能"
Const s_lang_0599 = "键入审核管理员帐号"
Const s_lang_0600 = "审核管理员列表"
Const s_lang_0601 = "键入被审核对象帐号"
Const s_lang_0602 = "被审核对象列表(支持通配符)"
Const s_lang_0603 = "键入例外帐号"
Const s_lang_0604 = "例外帐号列表(支持通配符)"
Const s_lang_0605 = "审核所有用户，除了列表内的例外帐号"
Const s_lang_0606 = "只审核列表内的用户，其他帐号都不审核"
Const s_lang_0607 = "审批时间不得超过"
Const s_lang_0608 = "小时"
Const s_lang_0609 = "审批邮件累积不得超过"
Const s_lang_0610 = "封"
Const s_lang_0611 = "超过限制后，未审批邮件将"
Const s_lang_0612 = "自动拒绝"
Const s_lang_0613 = "自动发送"

Const s_lang_0614 = "密码"
Const s_lang_0615 = "只读邮箱设置"
Const s_lang_0616 = "设定部分邮箱为只读邮箱"
Const s_lang_0617 = "启用只读邮箱功能"
Const s_lang_0618 = "只读邮箱帐号管理"
Const s_lang_0619 = "用户名"

Const s_lang_0620 = "主题关键字过滤"
Const s_lang_0621 = "处理方式: 删除邮件"
Const s_lang_0622 = "处理方式: 将邮件移到垃圾箱"
Const s_lang_0623 = "系统将过滤主题<font color=""#901111"">等于</font>(不区分大小写)指定关键字内容的邮件 (注意: 需启用""系统设置""中的相应功能)."
Const s_lang_0624 = "统一修改用户配置"
Const s_lang_0625 = "对所有用户的邮箱配置选项进行统一修改"
%>

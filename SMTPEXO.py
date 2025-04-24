import smtplib
import base64
from email.mime.text import MIMEText
import msal

# 配置信息
CLIENT_ID = 'e502ecd1-xxxxxxx'  # Azure应用注册的客户端ID Azure Enterprise Application, Application ID
CLIENT_SECRET = 'xxxxxxxS'  # Azure应用的客户端机密
TENANT_ID = '5axxxxxx-xxx-xxx'  # Azure目录（租户）ID
FROM_EMAIL = 'user02@yaba.top'  # 必须与Azure应用权限匹配的邮箱
TO_EMAIL = 'hello@hello.com'  # 收件人邮箱

# 获取OAuth访问令牌
def get_access_token():
    authority = f'https://login.microsoftonline.com/{TENANT_ID}'
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    # 请求作用域需对应Exchange Online
    result = app.acquire_token_for_client(scopes=['https://outlook.office365.com/.default'])
    return result.get('access_token')

# 创建邮件内容
msg = MIMEText('这是一封通过OAuth认证发送的测试邮件。')
msg['Subject'] = 'Python OAuth邮件测试'
msg['From'] = FROM_EMAIL
msg['To'] = TO_EMAIL

# 获取访问令牌
access_token = get_access_token()
if not access_token:
    raise ValueError("无法获取访问令牌，请检查应用配置")

# 构造XOAUTH2认证字符串
auth_string = f'user={FROM_EMAIL}\x01auth=Bearer {access_token}\x01\x01'
b64_auth_string = base64.b64encode(auth_string.encode()).decode()

# 连接SMTP服务器并发送邮件
try:
    with smtplib.SMTP('smtp.office365.com', 587) as server:
        server.set_debuglevel(2)  # Enable Debug mode
        ehlo_host = "UAT-EXCH-C01" # Hostname need follow some standard otherwise,5.5.4 Invalid domain name will be return
        server.ehlo(ehlo_host)
        server.starttls()  # 启用TLS加密
        server.ehlo(ehlo_host)
        # 发送XOAUTH2认证请求
        code, response = server.docmd('AUTH XOAUTH2', b64_auth_string)
        if code != 235:
            raise smtplib.SMTPAuthenticationError(code, response)
        # 发送邮件
        server.sendmail(FROM_EMAIL, [TO_EMAIL], msg.as_string())
        print("邮件发送成功！")
except Exception as e:
    print(f"发送邮件时出现错误: {e}")

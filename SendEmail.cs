/// <summary>
/// 寄信功能(含附件圖片),
/// Author : heng,
/// Create Time : 2021/08 
/// </summary>
/// <param name="emailModel">信件內容model</param>
/// <param name="sendMailSetting">寄件者設定</param>

public void SendEmail(EmailModel emailModel,JObject sendMailSetting){
    try{
        var hostUrl         = sendMailSetting["mailServer"].ToString();         //smtp url
        var port            = sendMailSetting["mailServerPort"].ToInt32();      //smtp port
        var message         = new MailMessage();            //建立信件寄件者 收件者 標題 主旨
        message.From        = new MailAddress(sendMailSetting["ac"].ToString(),sendMailSetting["emailTitle"].ToString()); //寄件人信箱與標題
        message.To.Add(emailModel.userEmail);               //收件者信箱
        message.Subject     = emailModel.subject;           //主旨
        message.Body        = emailModel.htmlText;          //內容
        message.IsBodyHtml  = true;
        
        //附件圖片
        AlternateView htmlView = 
            AlternateView.CreateAlternateViewFromString(emailModel.htmlText, null, "text/html");

        LinkedResource imageResource = new LinkedResource(emailModel.filePath, "image/jpg");
        string[] fileNameSpilt = emailModel.filePath.Split(@"\");
        string fileName = fileNameSpilt[fileNameSpilt.Length - 1].Split(@".")[0];
        imageResource.ContentId         = fileName;
        imageResource.TransferEncoding  = TransferEncoding.Base64;
        htmlView.LinkedResources.Add(imageResource);
        message.AlternateViews.Add(htmlView);

        using (var client = new SmtpClient (hostUrl,port)) {
            client.Credentials = new NetworkCredential(sendMailSetting["ac"].ToString(), AesDecrypt(sendMailSetting["pwd"].ToString())); //寄信用的信箱帳號密碼
            client.EnableSsl = true;
            client.Send(message);
            client.Dispose();
        } 
    }
    catch (Exception e){
        Console.WriteLine($"error:{e.ToString()}");
        throw e;
    }
}

/// <summary>
/// EmailModel 
/// <summary>
public class EmailModel
{
        public string userName { get; set; }
        public string userEmail { get; set; }
        public string subject { get; set; }
        public string title { get; set; }
        public string htmlText { get; set; }
        public string filePath { get; set; }
}

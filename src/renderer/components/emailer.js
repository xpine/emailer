import nodemailer from 'nodemailer';
export const createTransporter = (user,pass) => { 
  return nodemailer.createTransport({
    host: 'smtp.exmail.qq.com',
    port: 465, // SMTP 端口
    secure: true, // 使用了 SSL
    auth: {
      user: user,
      // 这里密码不是qq密码，是你设置的smtp授权码
      pass: pass,
    }
  });
}

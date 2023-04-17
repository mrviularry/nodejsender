const fs = require('fs');
const nodemailer = require('nodemailer');
const HttpProxyAgent = require('http-proxy-agent');
const { CommunicationIdentityClient } = require('@azure/communication-identity');
const { EmailSender } = require('@azure/communication-email');
const { DefaultAzureCredential } = require('@azure/identity');
const htmlToImage = require('html-to-image');
const pdf = require('html-pdf');

// Read the configuration options from config.json
const config = JSON.parse(fs.readFileSync('config.json', 'utf-8'));

// Setup Azure email sender
let emailSender;
if (config.useAzure) {
  const endpoint = config.azure.endpoint;
  const credential = new DefaultAzureCredential();
  const identityClient = new CommunicationIdentityClient(endpoint, credential);
  emailSender = new EmailSender(identityClient);
}

// Read the recipients, SMTPs, and subjects from the text files
const recipients = fs.readFileSync(config.recipientsFile, 'utf-8').split('\n').filter((email) => email.trim() !== '');
const smtps = fs.readFileSync(config.smtpsFile, 'utf-8').split('\n').filter((smtp) => smtp.trim() !== '');
const subjects = fs.readFileSync(config.subjectsFile, 'utf-8').split('\n').filter((subject) => subject.trim() !== '');

// Read the email HTML body and attachment
const message = fs.readFileSync(config.messageFile, 'utf-8');
const attachmentHtml = config.convertAttachments ? fs.readFileSync(config.attachmentHtmlFile, 'utf-8') : '';

// Set sleep time in seconds between each email and between threads
const sleepTime = config.sleepTime;
const threadPause = config.threadPause;

function extractDomainName(email) {
  const domain = email.split('@')[1].split('.')[0];
  return domain.charAt(0).toUpperCase() + domain.slice(1);
}

function createTransporter(smtpConfig) {
  return nodemailer.createTransport({
    ...smtpConfig,
    agent: config.useProxy ? new HttpProxyAgent(config.proxy) : null,
  });
}

async function sendEmail(recipientEmail, smtpConfig, subject, message, attachmentHtml) {
  const domainName = extractDomainName(recipientEmail);

  const personalizedSubject = subject.replace('{domainName}', domainName).replace('{email}', recipientEmail);
  const personalizedMessage = message.replace(/{domainName}/g, domainName).replace(/{email}/g, recipientEmail);
  const personalizedFromName = config.email.fromName.replace('{domainName}', domainName).replace('{email}', recipientEmail);

  const mailOptions = {
    from: `"${personalizedFromName}" <${config.email.fromEmail}>`,
    to: recipientEmail,
    subject: personalizedSubject,
    html: personalizedMessage,
    headers: config.customHeaders,
  };

  if (config.convertAttachments) {
    const imageBuffer = await htmlToImage.toPng(attachmentHtml);    
    const pdfBuffer = await new Promise((resolve, reject) => {
      pdf.create(attachmentHtml).toBuffer((err, buffer) => {
        if (err) reject(err);
        resolve(buffer);
      });
    });

    mailOptions.attachments = [
      {
        filename: 'attachment.png',
        content: imageBuffer.toString('base64'),
        encoding: 'base64',
      },
      {
        filename: 'attachment.pdf',
        content: pdfBuffer.toString('base64'),
        encoding: 'base64',
      },
    ];
  }

  if (config.useAzure) {
    const messageRequest = {
      from: mailOptions.from,
      to: [mailOptions.to],
      subject: mailOptions.subject,
      html: mailOptions.html,
      customHeaders: mailOptions.headers,
      attachments: mailOptions.attachments,
    };

    try {
      const messageId = await emailSender.sendMessage(messageRequest);
      console.log(`Email sent to ${recipientEmail}:`, messageId);
    } catch (error) {
      console.log(`Error sending email to ${recipientEmail}:`, error);
    }

  } else {
    const transporter = createTransporter(smtpConfig);
    transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
        console.log(`Error sending email to ${recipientEmail}:`, error);
      } else {
        console.log(`Email sent to ${recipientEmail}:`, info.response);
      }
    });
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

(async () => {
  const totalRecipients = recipients.length;
  const halfRecipients = Math.ceil(totalRecipients / 2);

  const thread1Recipients = recipients.slice(0, halfRecipients);
  const thread2Recipients = recipients.slice(halfRecipients);

  const thread1 = async () => {
    for (let i = 0; i < thread1Recipients.length; i++) {
      const recipientEmail = thread1Recipients[i];
      const smtpConfig = JSON.parse(smtps[i % smtps.length]);
      const subject = subjects[i % subjects.length];
      await sendEmail(recipientEmail, smtpConfig, subject, message, attachmentHtml);
      await sleep(sleepTime * 1000);
    }
  };

  const thread2 = async () => {
    await sleep(threadPause * 1000);
    for (let i = 0; i < thread2Recipients.length; i++) {
      const recipientEmail = thread2Recipients[i];
      const smtpConfig = JSON.parse(smtps[(i + halfRecipients) % smtps.length]);
      const subject = subjects[(i + halfRecipients) % subjects.length];
      await sendEmail(recipientEmail, smtpConfig, subject, message, attachmentHtml);
      await sleep(sleepTime * 1000);
    }
  };

  await Promise.all([thread1(), thread2()]);
})();

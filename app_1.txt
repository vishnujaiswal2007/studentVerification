import express from "express";
import { google } from "googleapis";
import { MongoClient } from "mongodb";
import fs from "fs";
import dotenv from "dotenv";
dotenv.config();
const URL = process.env.Data_URL;

const app = express();
app.use(express.json());
const PORT = process.env.PORT;

//Autherization Process
const getAuthClient = async () => {
  const credentials = JSON.parse(fs.readFileSync("credentials.json", "utf-8"));
  const { client_secret, client_id, redirect_uris } = credentials.web;
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  );
  const token = JSON.parse(fs.readFileSync("token.json", "utf-8"));
  oAuth2Client.setCredentials(token);

  return google.gmail({ version: "v1", auth: oAuth2Client });
};

const watchInbox = async (gmail) => {
  // const res = await gmail.users.watch({
  await gmail.users.watch({
    userId: "me",
    requestBody: {
      labelIds: ["INBOX"],
      topicName: "projects/studentverification-460010/topics/verification",
    },
  });
  //   console.log("📡 Watch response:", res.data);
};

const parseVerificationRequest = (bodyLines) => {
  const [PRG = "", YR = "", EN = "", RN = "", GT = ""] = bodyLines.map((line) =>
    (line || "").trim().toUpperCase()
  );
  const allFields = [PRG, YR, EN, RN, GT];
  const isValid = (val) => val !== undefined && val !== null && val !== "";

  return allFields.every(isValid) ? { PRG, YR, EN, RN, GT } : null;
};

const formatMessageByCourse = (courseCode, res) => {
  switch (courseCode) {
    case "BCOM3":
return `
    <p>The record of ${res.NM} availiable with ACC is as follows:</p>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse;">
      <tr><th colspan="2">${res.CC},${res.YR}</th></tr>
      <tr><td><strong>Name:</strong> ${res.NM}</td><td><strong>En No.:</strong> ${res.EN}</td></tr>
      <tr><td><strong>Father Name:</strong> ${res.GN}</td><td><strong>Roll No.:</strong> ${res.RN}</td></tr>
      <tr><td colspan="2"><strong>Mother Name:</strong> ${res.MN}</td></tr>
      <tr><th>Subject</th><th>Marks Obtained</th></tr>
      <tr><td>${res.SB1}</td><td>${res.M1}</td></tr>
      <tr><td>${res.SB2}</td><td>${res.M2}</td></tr>
      <tr><td>${res.SB3}</td><td>${res.M3}</td></tr>
      <tr><td>${res.SB4}</td><td>${res.M4}</td></tr>
      <tr><td>${res.SB5}</td><td>${res.M5}</td></tr>
      <tr><td>${res.SB6}</td><td>${res.M6}</td></tr>
      <tr><td><strong>Part-I:</strong> ${res.PM1}</td><td><strong>Part-II:</strong> ${res.PM2}</td></tr>
      <tr><td><strong>Total:</strong> ${res.TOT}</td><td><strong>Result:</strong> ${res.RES}</td></tr>
      <tr><td colspan="2"><strong>Grand Total:</strong> ${res.GT}</td></tr>
      <tr><td colspan="2"><strong>Remarks:</strong> ${res.SE}</td></tr>
    </table>
     <p>This is for your kind information and further action at your end.</p>
    <p>Thank you,<br>Support ACC Team</p>
  `;

    case "BALLBH":
      return `
      <p>The record of ${res.NM} availiable with ACC is as follows:</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse;">
      <tr><th colspan="3">${res.MRK_HEAD}, ${res.YR}</th></tr>
      <tr><td colspan="2"><strong>Name:</strong> ${res.NM}</td><td colspan="2"><strong>En No.:</strong> ${res.EN}</td></tr>
      <tr><td colspan="2"><strong>Father Name:</strong> ${res.GN}</td><td colspan="2"><strong>Roll No.:</strong> ${res.RN}</td></tr>
      <tr><td colspan="2"><strong>Mother Name:</strong> ${res.MN}</td></tr>
      <tr><th style="width: 70px;">Paper</th><th>Subject</th><th>Marks Obtained</th></tr> 
      <tr><td style="width: 70px;">1</td><td>${res.S1}</td><td style="text-align: right";>${res.T1}</td></tr>
      <tr><td style="width: 70px;">2</td><td>${res.S2}</td><td style="text-align: right";>${res.T2}</td></tr>
      <tr><td style="width: 70px;">3</td><td>${res.S3}</td><td style="text-align: right";>${res.T3}</td></tr>
      <tr><td style="width: 70px;">4</td><td>${res.S4}</td><td style="text-align: right";>${res.T4}</td></tr>
      <tr><td style="width: 70px;">5</td><td>${res.S5}</td><td style="text-align: right";>${res.T5}</td></tr>
      <tr><td style="width: 70px;">6</td><td>${res.S6}</td><td style="text-align: right";>${res.T6}</td></tr>
      <tr><td colspan="2" style="text-align: right";><strong>Sem Total :</strong></td><td style="text-align: right";><strong>${res.TOT}</strong></td></tr>
      <tr>
      <td colspan="3" >
      <strong>Sem-I:</strong>
      (${res.SEM_1})
      <span style="margin-left: 50px;">
      <strong>Sem-II:</strong>
      (${res.SEM_2})
       <span style="margin-left: 50px;">
       <strong>Sem-III:</strong>
      (${res.SEM_3})
      <span style="margin-left: 50px;">
      <strong>Sem-IV:</strong>
      (${res.SEM_4})
      <span style="margin-left: 50px;">
      <strong>Sem-V:</strong>
      (${res.SEM_5})
      </td>
      </tr>
      <tr>
      <td colspan="3" >
      <strong>Sem-VI:</strong>
      (${res.SEM_6})
      <span style="margin-left: 50px;">
      <strong>Sem-VII:</strong>
      (${res.SEM_7})
       <span style="margin-left: 50px;">
       <strong>Sem-VIII:</strong>
      (${res.SEM_8})
      <span style="margin-left: 50px;">
      <strong>Sem-IX:</strong>
      (${res.SEM_9})
      <span style="margin-left: 50px;">
      <strong>Sem-X:</strong>
      (${res.TOT})
      </td>
      </tr>
      <tr><td colspan="3"><strong>Result:</strong> ${res.RES} <span style="margin-left: 350px;"><strong>Grand Total : </strong>${res.GT}</td></tr>
      <tr><td colspan="3"><strong>Remarks:</strong> ${res.SE}</tr>
      </table>
       <p>This is for your kind information and further action at your end.</p>

Thank you, for email and 
Support ACC Team`;

    default:
      return `
      <p>The record of ${res.NM} availiable with ACC is as follows:</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse;">
      <tr><th colspan="3">${res.CC}, ${res.YR}</th></tr>
      <tr><td colspan="2"><strong>Name:</strong> ${res.NM}</td><td colspan="2"><strong>En No.:</strong> ${res.EN}</td></tr>
      <tr><td colspan="2"><strong>Father Name:</strong> ${res.GN}</td><td colspan="2"><strong>Roll No.:</strong> ${res.RN}</td></tr>
      <tr><td colspan="2"><strong>Mother Name:</strong> ${res.MN}</td></tr>
      <tr><th style="width: 70px;">Paper</th><th>Subject</th><th>Marks Obtained</th></tr> 
      <tr><td style="width: 70px;">1</td><td>${res.SB1}</td><td style="text-align: right";>${res.SUB1}</td></tr>
      <tr><td style="width: 70px;">2</td><td>${res.SB2}</td><td style="text-align: right";>${res.SUB2}</td></tr>
      <tr>
      <td colspan="3" >
      <strong>Part-I:</strong>
      (${res.PM1})
      <span style="margin-left: 150px;">
      <strong>Part-II:</strong>
      (${res.PM2})
      <span style="margin-left: 150px;">
      <strong>Total:</strong>
      (${res.TOT})
      </td>
      </tr>
      <tr><td colspan="3"><strong>Result:</strong> ${res.RES} <span style="margin-left: 350px;"><strong>Grand Total : </strong>${res.GT}</td></tr>
      <tr><td colspan="3"><strong>Remarks:</strong> ${res.SE}</tr>
      </table>
       <p>This is for your kind information and further action at your end.</p>

Thank you, for email and 
Support ACC Team`;
  }
};

const sendReplyEmail = async (
  gmail,
  { from, subject, messageId, threadId, body, isHtml=true}
) => {
  const replySubject = subject.startsWith("Re:") ? subject : `Re: ${subject}`;
  const contentType = isHtml ? "text/html" : "text/plain";
  
  const mimeMessage =
    `To: ${from}\r\n` +
    `Subject: ${replySubject}\r\n` +
    `In-Reply-To: ${messageId}\r\n` +
    `References: ${messageId}\r\n` +
    `Content-Type: ${contentType}; charset="UTF-8"\r\n\r\n` +
    body;

  const encodedMessage = Buffer.from(mimeMessage)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");

  await gmail.users.messages.send({
    userId: "me",
    requestBody: {
      raw: encodedMessage,
      threadId,
    },
  });
};

const listmessages = async (gmail) => {
  const res = await gmail.users.messages.list({
    userId: "me",
    q: "is:unread",
    maxResults: 1,
  });

  const messages = res.data.messages;
  if (!messages || messages.length === 0) {
    console.log("📭 No new unread messages.");
    return;
  } else {
    console.log(`📨 Found ${messages.length} unread message(s):`);
    for (const msg of messages) {
      const msgRes = await gmail.users.messages.get({
        userId: "me",
        id: msg.id,
        format: "Full",
      });

      const headers = msgRes.data.payload.headers || [];

      const from =
        headers.find((h) => h.name === "From")?.value || "(Unknown Sender";

      const headerMap = Object.fromEntries(
        headers.map((h) => [h.name, h.value])
      );

      const messageIdHeader = headers.find(
        (h) => h.name === "Message-ID"
      )?.value;

      const subject =
        headers.find((h) => h.name === "Subject")?.value || "(No Subject)";

      // Mark the message as read by removing the "UNREAD" label
      try {
        await gmail.users.messages.modify({
          userId: "me",
          id: msg.id,
          resource: {
            removeLabelIds: ["UNREAD"],
          },
        });
        console.log(`✅ Marked message as read: ${subject}`);
      } catch (err) {
        console.error(`❌ Failed to mark message as read: ${subject}`, err);
      }

      // Reply to those whose send reply on same email thread
      if ("In-Reply-To" in headerMap || "References" in headerMap) {
        const body = `Every verification need a fresh email in the follwing format:
         Subject: VERIFICATION
      Programme code,
      Year 
      Enrolment Number,
      Roll Number,
      Grand Total
Note: 
1:- All fields are required.
2:- Verification through e-mail is only for BA, B.Sc. and B.Com Final Year.
3:- The Programme Code for:
     BA           : PRA262
     B.Sc         : PRA270
     B.Com        : PRA351
     BALLB(Hons.) : PRC265
4:- It is requested not to add Signature at the end of the email.

Thank you,
Support ACC Team`
        await sendReplyEmail(gmail, {
          from,
          subject,
          messageId: messageIdHeader,
          threadId: msgRes.data.threadId,
          body,
          isHtml:false
        });
        console.log("Skipping reply email.");
        continue;
      }

      // console.log("from", from);

      if (from === "ACC Verification <verification_counter@allduniv.ac.in>") {
        //get the text body
        let body = "";
        const parts = msgRes.data.payload.parts || [msgRes.data.payload];

        for (const part of parts) {
          if (part.mimeType === "text/plain" && part.body?.data) {
            body = Buffer.from(part.body.data, "base64").toString("utf-8");
            break;
          }
        }

        const bodyPreview = body.slice(0, 100).trim();
        const bodyLines = bodyPreview
          .split("\n")
          .filter((line) => line.trim() !== "");

        const parsed = parseVerificationRequest(bodyLines);

        if (subject === "VERIFICATION" && parsed) {
          const { PRG, YR, EN, RN, GT } = parsed;
          const client = new MongoClient(URL);

          try {
            await client.connect();
            const db = client.db("COURSES");
            const CL = await db.collection("FINAL").findOne({ PRG_CODE: PRG });

            if (!CL) {
              const body =
  `Verification can be done only for Programme(s):
     BA           : PRA262
     B.Sc         : PRA270
     B.Com        : PRA351
     BALLB(Hons.) : PRC265.\n\nThank you,\nSupport ACC Team`;

              await sendReplyEmail(gmail, {
                from,
                subject,
                messageId: messageIdHeader,
                threadId: msgRes.data.threadId,
                body,
                isHtml:false
              });
            } else {
              const studentDb = client.db(CL.COURSE);
              const student = await studentDb
                .collection(CL.DB_CL)
                .findOne({ YR, EN, RN, GT, PDF: "PDF" });

              const body = student
                ? formatMessageByCourse(CL.DB_CL, student)
                : "<p>NO RECORD FOUND of the given data.\n\nThank you,\nSupport ACC Team</p>";

              await sendReplyEmail(gmail, {
                from,
                subject,
                messageId: messageIdHeader,
                threadId: msgRes.data.threadId,
                body,
                isHtml:true
              });
            }
          } catch (err) {
            console.error("Error:", err);
          } finally {
            await client.close();
          }
        } else {
          const instructions = `The data for Verification should be in following format:
    Subject: VERIFICATION
      Programme code,
      Year 
      Enrolment Number,
      Roll Number,
      Grand Total
Note: 
1:- All fields are required.
2:- Verification through e-mail is only for BA, B.Sc. and B.Com Final Year.
3:- The Programme Code for:
     BA           : PRA262
     B.Sc         : PRA270
     B.Com        : PRA351
     BALLB(Hons.) : PRC265
4:- Please send fresh email with Subject: VERIFICATION and details in capital letter and in sequence as detailed.
5:- It is requested not to add Signature at the end of the email.

Thank you,
Support ACC Team`;
          await sendReplyEmail(gmail, {
            from,
            subject,
            messageId: messageIdHeader,
            threadId: msgRes.data.threadId,
            body: instructions,
            isHtml:false
          });
        }
      } else {
        const warning = `You are an unauthorized user to use it. It is a system generated mail, please don't reply.  
Thank you,
Support ACC Team`;
        await sendReplyEmail(gmail, {
          from,
          subject,
          messageId: messageIdHeader,
          threadId: msgRes.data.threadId,
          body: warning,
          isHtml:false
        });
      }
    }
  }
};

app.post("/", async (req, res) => {
  // console.log('Pub/Sub Notification Received:', req.body);
  const message = req.body.message;
  if (!message?.data) {
    return res.status(400).send("Bad Request: Missing message data");
  } else {
    try {
      const gmail = await getAuthClient();
      await listmessages(gmail);
      res.status(200).send("OK"); // Acknowledge message
    } catch (error) {
      console.error("Error handling webhook:", error);
      res.status(500).send("Internal Server Error");
    }
  }
});

var server = app.listen(PORT, async () => {
  try {
    const gmail = await getAuthClient();
    await watchInbox(gmail);
    const host = server.address().address;
    const port = server.address().port;
    console.log(`Server is running on http://${host}:${port}`);
  } catch (error) {
    console.log("Error in Server running", error);
  }
});

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
      return `The provide data is matched with record. However please co-relate with details :

                       ${res.CC}
--------------------------------------------------------------------------------
Name: ${res.NM}                            En No.: ${res.EN}
Father Name : ${res.GN}                    Roll No.  : ${res.RN}
Mother Name : ${res.MN}
---------------------------------------------------------------------------------
Subject                                                 Marks Obtain       
---------------------------------------------------------------------------------
${res.SB1}                         ${res.M1}
${res.SB2}                                          ${res.M2}
${res.SB3}                        ${res.M3}
${res.SB4}                                     ${res.M4}
${res.SB5}        ${res.M5}
${res.SB6}          ${res.M6}
---------------------------------------------------------------------------------
Part-I:(${res.PM1})         Part-II:(${res.PM2})           Total : ${res.TOT}
---------------------------------------------------------------------------------
Result  : ${res.RES}               Grand Total : ${res.GT}
${res.SE}

Thank you, for email and 
Support ACC Team`;

    case "BALLBH":
      return `The provide data is matched with record. However please co-relate with details :

                       ${res.MRK_HEAD}
                             ${res.MRK_HEAD_2}, ${res.YR}
--------------------------------------------------------------------------------
Name: ${res.NM}                             Roll No.  : ${res.RN}                     
Father Name : ${res.GN}                     En No.: ${res.EN}                  
Mother Name : ${res.MN}                    
---------------------------------------------------------------------------------
Subject                                                 Marks Obtain       
---------------------------------------------------------------------------------
${res.S1}                         ${res.T1}
${res.S2}                                          ${res.T2}
${res.S3}                        ${res.T3}
${res.S4}                                     ${res.T4}
${res.S5}        ${res.T5}
${res.S6}          ${res.T6}
---------------------------------------------------------------------------------
                             Xth Sem Total : ${res.TOT}
                             ---------------------
Sem-I:(${res.SEM_1})   Sem-II:(${res.SEM_2})   Sem-III:(${res.SEM_3})
Sem-IV:(${res.SEM_4})   Sem-V:(${res.SEM_5})   Sem-VI:(${res.SEM_6})
Sem-VII:(${res.SEM_7})   Sem-VIII:(${res.SEM_8})   Sem-IX:(${res.SEM_9})
---------------------------------------------------------------------------------
Result  : ${res.RES}               Grand Total : ${res.GT}
${res.SE}

Thank you, for email and 
Support ACC Team`;

    default:
      return `The provide data is matched with record. However please co-relate with details :

                           ${res.CC}
--------------------------------------------------------------------------------
Name: ${res.NM}                            En No.: ${res.EN}
Father Name : ${res.GN}                    Roll No.  : ${res.RN}
Mother Name : ${res.MN}
---------------------------------------------------------------------------------
Subject                                        Marks Obtain
---------------------------------------------------------------------------------
${res.SB1}                                         ${res.SUB1}
${res.SB2}                                         ${res.SUB2}
---------------------------------------------------------------------------------
Part-I:(${res.PM1})         Part-II:(${res.PM2})           Total : ${res.TOT}                               
---------------------------------------------------------------------------------
Result  : ${res.RES}              Grand Total : ${res.GT}
${res.SE}

Thank you, for email and 
Support ACC Team`;
  }
};

const sendReplyEmail = async (
  gmail,
  { from, subject, messageId, threadId, body }
) => {
  const replySubject = subject.startsWith("Re:") ? subject : `Re: ${subject}`;

  const mimeMessage =
    `To: ${from}\r\n` +
    `Subject: ${replySubject}\r\n` +
    `In-Reply-To: ${messageId}\r\n` +
    `References: ${messageId}\r\n` +
    `Content-Type: text/plain; charset="UTF-8"\r\n\r\n` +
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
              });
            } else {
              const studentDb = client.db(CL.COURSE);
              const student = await studentDb
                .collection(CL.DB_CL)
                .findOne({ YR, EN, RN, GT, PDF: "PDF" });

              const body = student
                ? formatMessageByCourse(CL.DB_CL, student)
                : "NO RECORD FOUND of the given data.\n\nThank you,\nSupport ACC Team";

              await sendReplyEmail(gmail, {
                from,
                subject,
                messageId: messageIdHeader,
                threadId: msgRes.data.threadId,
                body,
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

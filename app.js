import express from "express";
import { google } from "googleapis";
import { MongoClient } from "mongodb";
import xlsx from "xlsx";
import path from "path";
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


const parseVerificationRequestExcel = (rawRow) => {
  const fieldOrder = [
    "PRG_CODE",
    "YEAR",
    "ENROLMENT",
    "ROL_NUMBER",
    "NAME",
    "FATHER_NAME",
    "MOTHER_NAME",
    "GRAND_TOTAL",
  ];

  const rowArray = fieldOrder.map((key) =>
    (rawRow[key] || "").toString().trim().toUpperCase()
  );

  const [
    PRG_CODE,
    YEAR,
    ENROLMENT,
    ROL_NUMBER,
    NAME,
    FATHER_NAME,
    MOTHER_NAME,
    GRAND_TOTAL,
  ] = rowArray;

  const allFields = rowArray;
  const isValid = (val) => val !== undefined && val !== null && val !== "";

  return allFields.every(isValid)
    ? {
        PRG_CODE,
        YEAR,
        ENROLMENT,
        ROL_NUMBER,
        NAME,
        FATHER_NAME,
        MOTHER_NAME,
        GRAND_TOTAL,
      }
    : null;
};





const formatMessageByCourse = (courseCode, res) => {
  switch (courseCode) {
    
      case "LLBH6":
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
      
      <tr><td colspan="3"><strong>Result:</strong> ${res.RES} <span style="margin-left: 350px;"><strong>Grand Total : </strong>${res.GT}</td></tr>
      <tr><td colspan="3"><strong>Remarks:</strong> ${res.SE}</tr>
      </table>
       <p>This is for your kind information and further action at your end.</p>

Thank you, for email and 
Support ACC Team`;
  
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
  { from, subject, messageId, threadId, body, isHtml = true }
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

const createMimeWithAttachment = (
  to,
  subject,
  threadId,
  messageId,
  bodyText,
  filename,
  base64Content
) => {
  const boundary = "__BOUNDARY__";
  // const filename = path.basename(filePath);
  // const fileContent = fs.readFileSync(filePath).toString("base64");

  const messageParts = [
    `To: ${to}`,
    `Subject: ${subject}`,
    `In-Reply-To: ${messageId}`,
    `References: ${messageId}`,
    `Thread-Id: ${threadId}`,
    `MIME-Version: 1.0`,
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    ``,
    `--${boundary}`,
    `Content-Type: text/plain; charset="UTF-8"`,
    `Content-Transfer-Encoding: 7bit`,
    ``,
    bodyText,
    ``,
    `--${boundary}`,
    `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="${filename}"`,
    `Content-Disposition: attachment; filename="${filename}"`,
    `Content-Transfer-Encoding: base64`,
    ``,
    base64Content,
    `--${boundary}--`,
  ];

  const mimeMessage = messageParts.join("\r\n");

  return Buffer.from(mimeMessage)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
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

      if (from === "ACC Verification <verification_counter@allduniv.ac.in>") {
        if (subject === "VERIFICATION") {
          if ("In-Reply-To" in headerMap || "References" in headerMap) {
            const body = `Every verification need a fresh email.
          Thank you,
          Support ACC Team`;
            await sendReplyEmail(gmail, {
              from,
              subject,
              messageId: messageIdHeader,
              threadId: msgRes.data.threadId,
              body,
              isHtml: false,
            });
            console.log("Not Entertained");
            continue;
          } else {

            const parts = msgRes.data.payload.parts || [msgRes.data.payload];

            if (
              parts?.some(
                (part) =>
                  part.filename &&
                  part.filename.length > 0 &&
                  part.body?.attachmentId
              )
            ){
          
          
          
              //EXCEL: Reading and reply of excel file as attachment
            for (const part of parts || []) {
              if (
                part.filename &&
                part.filename.endsWith(".xlsx") &&
                part.body.attachmentId
              ) {
                const attachmentRes =
                  await gmail.users.messages.attachments.get({
                    userId: "me",
                    messageId: msg.id,
                    id: part.body.attachmentId,
                  });

                const dataBuffer = Buffer.from(
                  attachmentRes.data.data,
                  "base64"
                );
                const workbook = xlsx.read(dataBuffer, { type: "buffer" });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = xlsx.utils.sheet_to_json(worksheet);

                const client = new MongoClient(URL);
                await client.connect();

                const db = client.db("COURSES");
                const output = [];

                for (let rawRow of jsonData) {

                  let row = parseVerificationRequestExcel(rawRow)

                  let status = "Do verify from Administrative Computer Center (ACC)";

                  if (row) {

                  const {
                    PRG_CODE,
                    YEAR,
                    ROL_NUMBER,
                    ENROLMENT,
                    NAME,
                    FATHER_NAME,
                    MOTHER_NAME,
                    GRAND_TOTAL,
                  } = row;


                  const CL = await db
                    .collection("FINAL")
                    .findOne({ PRG_CODE: PRG_CODE });

                  
                  if (CL) {
                    const studentDb = client.db(CL.COURSE);
                    const qury = {
                      YR: `${YEAR}`,
                      RN: `${ROL_NUMBER}`,
                      EN: `${ENROLMENT}`,
                      NM: `${NAME}`,
                      GN: `${FATHER_NAME}`,
                      MN: `${MOTHER_NAME}`,
                      GT: `${GRAND_TOTAL}`,
                      PDF: "PDF",
                    };
                    const student = await studentDb
                      .collection(CL.DB_CL)
                      .findOne(qury);
                    if (student) {
                      status = "Record Matched";
                      row = {
                        ...row,
                        RESULT: student.RES,
                        SECOND_EXAMINATION: student.SE,
                        STATUS: status,
                      };

                      output.push(row);
                    } else {
                      row = {
                        ...row,
                        RESULT: "NA",
                        SECOND_EXAMINATION: "NA",
                        STATUS: status,
                      };

                      output.push(row);
                    }
                  } else {
                    row = {
                      ...row,
                      RESULT: "NA",
                      SECOND_EXAMINATION: "NA",
                      STATUS: status,
                      REMARK: "Program Code does not Exist",
                    };

                    output.push(row);
                  }
                  }else{

                    row = {
                      ...rawRow,
                      RESULT: "NA",
                      SECOND_EXAMINATION: "NA",
                      STATUS: status,
                      REMARK: "All fields are necessary",
                    };

                    output.push(row);
                  }
                }

                await client.close();
                const updatedSheet = xlsx.utils.json_to_sheet(output);
                const updatedWorkbook = xlsx.utils.book_new();
                xlsx.utils.book_append_sheet(
                  updatedWorkbook,
                  updatedSheet,
                  sheetName
                );
                // const updatedFilePath = `./updated_${Date.now()}.xlsx`;
                // xlsx.writeFile(updatedWorkbook, updatedFilePath);

                const fileBuffer = xlsx.write(updatedWorkbook, {
                  type: "buffer",
                  bookType: "xlsx",
                });
                const base64Excel = fileBuffer.toString("base64");

                // 📤 Send reply with attachment
                const fileName = `updated_${Date.now()}.xlsx`;
                const rawMime = createMimeWithAttachment(
                  from,
                  `Re: ${subject}`,
                  msgRes.data.threadId,
                  messageIdHeader,
                  "Please find the updated verification results in the attached Excel file.",
                  fileName,
                  base64Excel
                );

                await gmail.users.messages.send({
                  userId: "me",
                  requestBody: {
                    raw: rawMime,
                    threadId: msgRes.data.threadId,
                  },
                });

                console.log("✅ Replied with updated Excel file.");
                break; // stop after first matching Excel
              }
              }} 
              
              //verification by simple message
              
              else {
                let body = "";
                for (const part of parts) {
                  if (part.mimeType === "text/plain" && part.body?.data) {
                    body = Buffer.from(part.body.data, "base64").toString("utf-8");
                    break;
                  }
                }
                // console.log("Body Message is ", body)
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
                    const CL = await db
                      .collection("FINAL")
                      .findOne({ PRG_CODE: PRG });

                    if (!CL) {
                      const body = `Verification can be done only for Programme(s):
                     BA           : PRA262
                     B.Sc         : PRA270
                     B.Com        : PRA351
                     BALLB(Hons.) : PRC265
                     LLB(Hons.)   : PRC272
                     .\n\nThank you,\nSupport ACC Team`;

                      await sendReplyEmail(gmail, {
                        from,
                        subject,
                        messageId: messageIdHeader,
                        threadId: msgRes.data.threadId,
                        body,
                        isHtml: false,
                      });
                    } else {
                      const studentDb = client.db(CL.COURSE);
                      const student = await studentDb
                        .collection(CL.DB_CL)
                        .findOne({ YR, EN, RN, GT, PDF: "PDF" });

                      const body = student
                        ? formatMessageByCourse(CL.DB_CL, student)
                        : "<p>Do verify from Administrative Computer Center (ACC).\n\nThank you,\nSupport ACC Team</p>";

                      await sendReplyEmail(gmail, {
                        from,
                        subject,
                        messageId: messageIdHeader,
                        threadId: msgRes.data.threadId,
                        body,
                        isHtml: true,
                      });
                    }
                  } catch (err) {
                    console.error("Error:", err);
                  } finally {
                    await client.close();
                  }
                }
              }
          }
        } else {
          const instructions = `The data for Verification should be in following format:

        Subject: VERIFICATION
          Programme code,
          Year 
          Enrolment Number,
          Roll Number,
          Grand Total


    Note: (for Single Verification)
    1:- All fields are required.
    2:- Verification through e-mail is only for Final Year/Semster.
    3:- Please send fresh email with Subject: VERIFICATION.
    4:- All details should be in capital letter and in-sequence as detailed.
    5:- It is requested not to add Signature at the end of the email.


    Note:(For Bulk Verification)
    1:- Kindly download the attached Excel Sheet.
    2;- Fill all the details.
    3:- All details are necessary and in CAPITAL letter.
    4:- Details should be in prescribed EXCEL format.
    5:- Compose a new email with Subject: VERIFICATION.
    6:- Attach the Excel sheet (having details)


    The Programme Code are:
         BA                       : PRA262
         B.Sc                   : PRA270
         B.Com                : PRA351
         BALLB(Hons.)     : PRC265
         LLB(Hons.)         : PRC272
    
    Thank you,
    Support ACC Team`;

          const filePath = "./VERIFICATION.xlsx"; // your local Excel file
          const filename = path.basename(filePath);
          const fileContent = fs.readFileSync(filePath).toString("base64");

          const rawMime = createMimeWithAttachment(
            from,
            `Re: ${subject}`,
            msgRes.data.threadId,
            messageIdHeader,
            instructions,
            filename,
            fileContent
          );

          await gmail.users.messages.send({
            userId: "me",
            requestBody: {
              raw: rawMime,
              threadId: msgRes.data.threadId, // ensures it's part of the same conversation
            },
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
          isHtml: false,
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

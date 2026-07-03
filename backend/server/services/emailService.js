const nodemailer = require("nodemailer");

function getTransporter() {
  const host = process.env.SMTP_HOST;
  const port = Number(process.env.SMTP_PORT || 587);
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;

  if (!host || !user || !pass) {
    return null;
  }

  return nodemailer.createTransport({
    host,
    port,
    secure: String(process.env.SMTP_SECURE || "false").toLowerCase() === "true",
    auth: { user, pass },
  });
}

async function sendOtpEmail({ to, name, otp, purpose }) {
  const transporter = getTransporter();
  const from = process.env.MAIL_FROM || "TraceWrite <no-reply@tracewrite.local>";
  
  let subject, intro;
  if (purpose === "mfa_login") {
    subject = "Your TraceWrite MFA Code";
    intro = "Use this code to complete your login.";
  } else if (purpose === "password_reset") {
    subject = "Reset Your TraceWrite Password";
    intro = "Use this code to reset your password.";
  } else {
    subject = "Verify your TraceWrite email";
    intro = "Use this code to verify your email and activate your account.";
  }

  const text = `${intro}\n\nCode: ${otp}\n\nThis code expires in 10 minutes.`;

  if (!transporter) {
    console.log(`[OTP DEV MODE] ${purpose} code for ${to}: ${otp}`);
    return { simulated: true };
  }

  try {
    await transporter.sendMail({
      from,
      to,
      subject,
      text,
    });
    return { simulated: false };
  } catch (err) {
    console.error(`[OTP EMAIL ERROR] Failed to send ${purpose} email to ${to}:`, err.message || err);
    if (String(process.env.OTP_DEV_FALLBACK || "false").toLowerCase() === "true") {
      console.log(`[OTP DEV MODE] ${purpose} code for ${to}: ${otp}`);
      return { simulated: true, error: err };
    }
    throw err;
  }
}

module.exports = { sendOtpEmail };

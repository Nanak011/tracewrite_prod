let mfaChallengeToken = "";
let resetEmail = "";

document.addEventListener("DOMContentLoaded", () => {
  const loginForm = document.getElementById("loginForm");
  if (loginForm) {
    loginForm.addEventListener("submit", async (e) => {
      e.preventDefault();
    try {
      const result = await API.post("/api/auth/login", {
        email: document.getElementById("email").value,
        password: document.getElementById("password").value,
      });

      if (result.mfaRequired) {
        mfaChallengeToken = String(result.challengeToken || "");
        document.getElementById("loginForm").style.display = "none";
        document.getElementById("mfaForm").style.display = "block";
        document.getElementById("forgotPasswordLink").style.display = "none";
        showMessage("authMessage", "MFA OTP sent to your email.", "info");
        return;
      }

      window.redirect("/dashboard");
    } catch (err) {
      showMessage("authMessage", err.message, "error");
    }
    });
  }

  const mfaForm = document.getElementById("mfaForm");
  if (mfaForm) {
    mfaForm.addEventListener("submit", async (e) => {
      e.preventDefault();
    try {
      await API.post("/api/auth/verify-mfa", {
        challengeToken: mfaChallengeToken,
        otp: document.getElementById("mfaOtp").value,
      });
      window.redirect("/dashboard");
    } catch (err) {
      showMessage("authMessage", err.message, "error");
    }
    });
  }

  const forgotPasswordLinkBtn = document.getElementById("forgotPasswordLinkBtn");
  if (forgotPasswordLinkBtn) {
    forgotPasswordLinkBtn.addEventListener("click", (e) => {
      e.preventDefault();
      document.getElementById("loginForm").style.display = "none";
      document.getElementById("forgotPasswordForm").style.display = "block";
      document.getElementById("forgotPasswordLink").style.display = "none";
      document.getElementById("authMessage").textContent = "";
    });
  }

  const backToLoginBtn = document.getElementById("backToLoginBtn");
  if (backToLoginBtn) {
    backToLoginBtn.addEventListener("click", () => {
      document.getElementById("forgotPasswordForm").style.display = "none";
      document.getElementById("resetPasswordForm").style.display = "none";
      document.getElementById("loginForm").style.display = "block";
      document.getElementById("forgotPasswordLink").style.display = "block";
      document.getElementById("authMessage").textContent = "";
      resetEmail = "";
    });
  }

  const backToForgotBtn = document.getElementById("backToForgotBtn");
  if (backToForgotBtn) {
    backToForgotBtn.addEventListener("click", () => {
      document.getElementById("resetPasswordForm").style.display = "none";
      document.getElementById("forgotPasswordForm").style.display = "block";
      document.getElementById("authMessage").textContent = "";
    });
  }

  const forgotPasswordForm = document.getElementById("forgotPasswordForm");
  if (forgotPasswordForm) {
    forgotPasswordForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      try {
        const email = document.getElementById("forgotEmail").value;
        await API.post("/api/auth/forgot-password", { email });
        
        resetEmail = email;
        document.getElementById("resetEmailLabel").textContent = email;
        document.getElementById("forgotPasswordForm").style.display = "none";
        document.getElementById("resetPasswordForm").style.display = "block";
        showMessage("authMessage", "If your email exists, a reset code has been sent.", "info");
      } catch (err) {
        showMessage("authMessage", err.message, "error");
      }
    });
  }

  const resetPasswordForm = document.getElementById("resetPasswordForm");
  if (resetPasswordForm) {
    resetPasswordForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      try {
        await API.post("/api/auth/verify-reset-otp", {
          email: resetEmail,
          otp: document.getElementById("resetOtp").value,
          newPassword: document.getElementById("newPassword").value,
        });
        
        showMessage("authMessage", "Password reset successfully! You can now login.", "info");
        document.getElementById("resetPasswordForm").style.display = "none";
        document.getElementById("loginForm").style.display = "block";
        document.getElementById("forgotPasswordLink").style.display = "block";
        
        // Pre-fill email for convenience
        document.getElementById("email").value = resetEmail;
        resetEmail = "";
      } catch (err) {
        showMessage("authMessage", err.message, "error");
      }
    });
  }
});

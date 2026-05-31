/**
 * API layer – backend communication for the Email Assistant add-on.
 * Handles user lookup and organization verification.
 */
function getUserInfo(email, idToken) {
    try {
        if (!SERVER_DOMAIN) {
            Logger.log("ERROR: SERVER_DOMAIN is not configured in script properties");
            return { error: "Configuration error", message: "Server domain is not configured. Please check script properties." };
        }

        if (!idToken) {
            Logger.log("ERROR: Google ID token is missing");
            return { error: "Authentication error", message: "Failed to get Google ID token. Please try refreshing the add-on." };
        }

        Logger.log("User lookup request - Email: " + (email ? email.substring(0, 2) + "***@" + (email.indexOf("@") >= 0 ? email.substring(email.indexOf("@")) : "") : "null") + ", Server: " + SERVER_DOMAIN);
        Logger.log("ID Token preview: " + (idToken ? idToken.substring(0, 20) + "..." : "null"));

        var response = UrlFetchApp.fetch(
            SERVER_DOMAIN + "/api/organizations/by-user-email?email=" + encodeURIComponent(email),
            {
                muteHttpExceptions: true,
                headers: { Authorization: "Bearer " + idToken },
            }
        );
        var code = response.getResponseCode();
        var responseText = response.getContentText();

        Logger.log("User lookup response code: " + code + " for email: " + (email ? email.substring(0, 2) + "***@" + (email.indexOf("@") >= 0 ? email.substring(email.indexOf("@")) : "") : "null"));
        if (code !== 200 && code !== 404 && code !== 401 && code !== 500) {
            Logger.log("Response preview: " + (responseText ? responseText.substring(0, 80) + "..." : "empty"));
        }

        if (code === 200) {
            var responseData = JSON.parse(responseText);
            Logger.log("User lookup successful - keys: " + Object.keys(responseData).join(", "));
            return responseData;
        } else if (code === 404) {
            // User is authenticated but not registered in our system
            try {
                var errorData = JSON.parse(responseText);
                Logger.log("User not registered in system");
                return { error: "User not registered", message: errorData.message || "User not found in ReplAI - Email Assistant database" };
            } catch (parseErr) {
                Logger.log("Error parsing 404 response: " + parseErr);
                return { error: "User not registered", message: "User not found in ReplAI - Email Assistant database" };
            }
        } else if (code === 401) {
            // Authentication failed
            try {
                var errorData = JSON.parse(responseText);
                Logger.log("Authentication failed");
                return {
                    error: "Authentication failed",
                    message: errorData.message || errorData.error || "Invalid Google ID token. Please check server configuration."
                };
            } catch (parseErr) {
                Logger.log("Error parsing 401 response: " + parseErr);
                return { error: "Authentication failed", message: "Invalid Google ID token" };
            }
        } else if (code === 500) {
            // Server error
            try {
                var errorData = JSON.parse(responseText);
                Logger.log("Server error");
                return {
                    error: "Server error",
                    message: errorData.message || errorData.error || "Server configuration error. Please contact support."
                };
            } catch (parseErr) {
                Logger.log("Error parsing 500 response: " + parseErr);
                return { error: "Server error", message: "Internal server error" };
            }
        } else {
            Logger.log("User lookup failed for " + (email ? email.substring(0, 2) + "***@" + (email.indexOf("@") >= 0 ? email.substring(email.indexOf("@")) : "") : "null") + " - HTTP " + code);
            try {
                var errorData = JSON.parse(responseText);
                return {
                    error: "Request failed",
                    message: errorData.message || errorData.error || "Unexpected error (HTTP " + code + ")"
                };
            } catch (err) {
                return { error: "Request failed", message: "Unexpected error (HTTP " + code + ")" };
            }
        }
    } catch (err) {
        Logger.log("Exception in user lookup: " + err.toString());
        return { error: "Exception", message: "Error connecting to server. Please try again." };
    }
}

/**
 * Calls the backend universal sign-out endpoint. The backend verifies the
 * Google ID token, revokes the Google OAuth grant, and clears stored tokens.
 * Best-effort: returns { ok: bool, code: number } and never throws.
 */
function revokeBackendSession(idToken) {
    try {
        if (!SERVER_DOMAIN) {
            Logger.log("Sign-out: SERVER_DOMAIN not configured");
            return { ok: false, code: 0 };
        }
        if (!idToken) {
            Logger.log("Sign-out: no ID token available");
            return { ok: false, code: 0 };
        }

        var response = UrlFetchApp.fetch(
            SERVER_DOMAIN + "/api/users/addon/signout",
            {
                method: "post",
                contentType: "application/json",
                payload: "{}",
                muteHttpExceptions: true,
                headers: { Authorization: "Bearer " + idToken },
            }
        );
        var code = response.getResponseCode();
        Logger.log("Sign-out: backend response code " + code);
        return { ok: code === 200, code: code };
    } catch (err) {
        Logger.log("Sign-out: exception calling backend: " + err.toString());
        return { ok: false, code: 0 };
    }
}

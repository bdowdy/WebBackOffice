using Twilio;
using Twilio.Rest.Api.V2010.Account;

namespace MSaleWebServer.Services
{
    /// <summary>
    /// Thin wrapper around the Twilio SDK — mirrors the VB6 GL.SMSText() call.
    /// Credentials (API Key SID/secret + Account SID) come from configuration:
    /// user-secrets in Development, environment variables (Twilio__*) in Production.
    /// Nothing sensitive is stored in appsettings.json.
    /// </summary>
    public class SmsService
    {
        private readonly string _from;
        private readonly ILogger<SmsService> _logger;

        public SmsService(IConfiguration config, ILogger<SmsService> logger)
        {
            _logger = logger;
            var accountSid = config["Twilio:AccountSid"]   ?? "";
            var apiKeySid  = config["Twilio:ApiKeySid"]    ?? "";
            var apiSecret  = config["Twilio:ApiKeySecret"] ?? "";
            _from          = config["Twilio:FromNumber"]   ?? "";
            // Initialize with a Twilio API Key (SK…) + secret scoped to the account.
            // API keys can be revoked/rotated independently of the account Auth Token.
            TwilioClient.Init(apiKeySid, apiSecret, accountSid);
        }

        /// <summary>
        /// Sends an SMS. Returns the Twilio message SID on success, empty string on failure.
        /// Mirrors the VB6 GL.SMSText(body, phoneNo, "") call.
        /// </summary>
        public string Send(string toNumber, string body)
        {
            if (string.IsNullOrWhiteSpace(toNumber)) return "";

            try
            {
                string normalized = NormalizePhone(toNumber);
                if (string.IsNullOrWhiteSpace(normalized)) return "";

                var message = MessageResource.Create(
                    to:   new Twilio.Types.PhoneNumber(normalized),
                    from: new Twilio.Types.PhoneNumber(_from),
                    body: body
                );

                _logger.LogInformation("SMS sent to {To} — SID: {Sid}", normalized, message.Sid);
                return message.Sid ?? "";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "SMS send failed to {To}", toNumber);
                return "";
            }
        }

        private static string NormalizePhone(string raw)
        {
            string digits = new string(raw.Where(char.IsDigit).ToArray());
            return digits.Length switch
            {
                10 => "+1" + digits,
                11 when digits[0] == '1' => "+" + digits,
                _ => digits.Length > 0 ? "+" + digits : ""
            };
        }
    }
}

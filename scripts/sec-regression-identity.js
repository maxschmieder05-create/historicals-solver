const SEC_USER_AGENT_CONFIGURATION_MESSAGE =
  "SEC_USER_AGENT is required for live SEC regression scripts and must identify the application plus a real, monitored operator email address.";

class SecRegressionIdentityError extends Error {
  constructor(message = SEC_USER_AGENT_CONFIGURATION_MESSAGE) {
    super(message);
    this.name = "SecRegressionIdentityError";
  }
}

function requireSecRegressionHeaders(env = process.env) {
  const userAgent = env.SEC_USER_AGENT?.trim() ?? "";
  if (!validSecRegressionUserAgent(userAgent)) {
    throw new SecRegressionIdentityError(
      `${SEC_USER_AGENT_CONFIGURATION_MESSAGE} Placeholder, reserved-example, and no-reply addresses are not accepted; configure SEC_USER_AGENT before this script can make a network request.`
    );
  }
  return Object.freeze({ "User-Agent": userAgent });
}

function validSecRegressionUserAgent(value) {
  if (typeof value !== "string" || value.length < 10 || value.length > 250) return false;
  if (/[\u0000-\u001f\u007f]/.test(value)) return false;

  const emailMatch = value.match(/\b([A-Z0-9._%+-]+)@([A-Z0-9.-]+\.[A-Z]{2,})\b/i);
  if (!emailMatch) return false;

  const normalized = value.toLowerCase();
  const localPart = emailMatch[1].toLowerCase();
  const domain = emailMatch[2].toLowerCase();
  const identityText = value.replace(emailMatch[0], " ").replace(/[^a-z0-9]+/gi, "");

  if (identityText.length < 2) return false;
  if (/\b(?:placeholder|change[-_. ]?me|replace[-_. ]?me|your[-_. ]?(?:app|email|name))\b/i.test(value)) return false;
  if (/^(?:example\.(?:com|org|net)|localhost|invalid)$/.test(domain)) return false;
  if (/(?:^|\.)example\./.test(domain)) return false;
  if (/\.(?:example|invalid|localhost|test)$/.test(domain) || /(?:^|\.)your-domain\./.test(domain)) return false;
  if (/^(?:no[-_. ]?reply|do[-_. ]?not[-_. ]?reply|donotreply)(?:\+.*)?$/.test(localPart)) return false;
  if (/^(?:test|example|placeholder|your[-_. ]?(?:email|name)|fake|dummy|sample|user)(?:[+._-].*)?$/.test(localPart)) return false;
  if (/\b(?:contact|your[-_. ]?email|email)@example\.(?:com|org|net)\b/.test(normalized)) return false;

  return true;
}

module.exports = {
  SEC_USER_AGENT_CONFIGURATION_MESSAGE,
  SecRegressionIdentityError,
  requireSecRegressionHeaders,
  validSecRegressionUserAgent
};

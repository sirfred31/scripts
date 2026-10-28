// --- CONFIGURATION MATRIX ---

const BOT_CONFIGS = [
  {
    CHANNEL_ID: 'CNWTZ92GY',
    CHANNEL_NAME: '#tr-it',
    EMAIL_SOURCES: [
      'trleadsuptasklogistics@openaccessbpo.com',
      'trleadsuppikea@openaccessbpo.com'
    ]
  },
  {
    CHANNEL_ID: 'C0B8TL021LZ',
    CHANNEL_NAME: '#Test',
    EMAIL_SOURCES: [
      'DoNotReply@notification.fortinet.net'
    ]
  },
  {
    CHANNEL_ID: 'CARPLD22H',
    CHANNEL_NAME: '#patch-it-team',
    EMAIL_SOURCES: [
      'anagram2fa@openaccessbpo.com',
      'DoNotReply@notification.fortinet.net',
      'notification.fortinet.net'
    ]
  },
  {
    CHANNEL_ID: 'C02LV66ND55',
    CHANNEL_NAME: 'it-support-clusterchari',
    EMAIL_SOURCES: [
      'anagram2fa@openaccessbpo.com',
      'DoNotReply@notification.fortinet.net',
      'notification.fortinet.net',
      'itdavao@openaccessbpo.com',
      'wellnessworksvpn@openaccessbpo.com',
      'trleadsuptasklogistics@openaccessbpo.com',
      'trleadsuppikea@openaccessbpo.com'

    ]
  },
  {
    CHANNEL_ID: 'C05LUST1X4M',
    CHANNEL_NAME: '#luggagehero-vpn-team',
    EMAIL_SOURCES: [
      'luggageherovpntoken@openaccessbpo.com'
    ]
  }
];

const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
const SIGNING_SECRET = PropertiesService.getScriptProperties().getProperty('SLACK_SIGNING_SECRET');
const MAX_WAIT_TIME_MS = 5 * 60 * 1000;   // how long the bot waits for the OTP email
const OTP_LOOKBACK_MS  = 2 * 60 * 1000;   // also accept an OTP that arrived up to 2 min BEFORE the request
const OTP_FRESH_MS     = 3 * 60 * 1000;   // an OTP email must be no older than this
const STATE_CLEANUP_MS = 10 * 60 * 1000;  // finished requests are remembered this long (prevents duplicate replies)

function automatedOtpBot() {
  // If the previous run is still going when the trigger fires again, skip this run.
  // Two runs at the same time could both find the same OTP and send it twice.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    console.log("Previous run still in progress - skipping this run.");
    return;
  }

  try {
    for (let config of BOT_CONFIGS) {
      console.log(`--- Processing Channel Profile: ${config.CHANNEL_ID} ---`);
      try {
        processParallelWorkflow(config);
      } catch (globalExecutionError) {
        console.log(`Critical crash caught for ${config.CHANNEL_ID}: ` + globalExecutionError.toString());
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function processParallelWorkflow(config) {
  const cacheEngine = PropertiesService.getScriptProperties();
  const httpOptions = {
    method: "get",
    headers: { "Authorization": "Bearer " + SLACK_TOKEN },
    muteHttpExceptions: true
  };

  // States created in Phase 1 of this run are already handled there
  const newlyCreatedKeys = new Set();

  // PHASE 1: DISCOVER NEW REQUESTS
  const nowSec = Math.floor(Date.now() / 1000);
  const historyUrl = `https://slack.com/api/conversations.history?channel=${config.CHANNEL_ID}&oldest=${nowSec - 70}`;

  let historyData;
  try {
    historyData = JSON.parse(UrlFetchApp.fetch(historyUrl, httpOptions).getContentText());
  } catch (e) {
    console.log("Slack history fetch error: " + e.toString());
  }

  if (historyData && historyData.ok && historyData.messages) {
    for (let msg of historyData.messages.slice().sort((a, b) => parseFloat(a.ts) - parseFloat(b.ts))) {
      if (msg.bot_id || msg.subtype === "bot_message" || !msg.user || !msg.text) continue;

      const cleanText = msg.text.toLowerCase().trim();
      if (cleanText.includes("token") || cleanText.includes("code")) {
        const stateKey = `state_${config.CHANNEL_ID}_${msg.ts}`;

        if (cacheEngine.getProperty(stateKey) === null) {
          const state = {
            agent: msg.user,
            ts: msg.ts,
            startTime: Date.now(),
            step: "WAITING_FOR_OTP"
          };
          cacheEngine.setProperty(stateKey, JSON.stringify(state));
          newlyCreatedKeys.add(stateKey);
          console.log(`Registered tracking state for Agent <@${msg.user}> on thread ${msg.ts}`);

          // Agent may have connected already -> the OTP email may already be here
          const code = findFreshOtp(config, state);
          if (code) {
            state.step = "DONE"; // mark as sent first -> the code is never sent twice
            cacheEngine.setProperty(stateKey, JSON.stringify(state));
            sendSlackMessage(config.CHANNEL_ID, `<@${state.agent}>, here is your active OTP code: \`${code}\``, state.ts);
          } else {
            sendSlackMessage(config.CHANNEL_ID, `<@${msg.user}>, pa connect po`, msg.ts);
          }
        }
      }
    }
  }

  // PHASE 2: KEEP LOOKING FOR THE OTP OF EVERY WAITING REQUEST (oldest first)
  let allProperties;
  try {
    allProperties = cacheEngine.getProperties();
  } catch (e) {
    console.log("Cache storage engine unavailable: " + e.toString());
    return;
  }

  const prefix = `state_${config.CHANNEL_ID}_`;

  let activeKeys = [];
  for (let key in allProperties) {
    if (key.startsWith(prefix) && !newlyCreatedKeys.has(key)) {
      activeKeys.push({ key: key, ts: parseFloat(key.replace(prefix, "")) });
    }
  }
  activeKeys.sort((a, b) => a.ts - b.ts);

  for (let item of activeKeys) {
    let state = JSON.parse(allProperties[item.key]);
    const age = Date.now() - state.startTime;

    // Finished requests stay for a while so the same Slack message is not picked up twice
    if (state.step === "DONE") {
      if (age > STATE_CLEANUP_MS) cacheEngine.deleteProperty(item.key);
      continue;
    }

    // No confirmation needed: whether the agent connected before or after
    // typing 'token', the moment the OTP email is in Gmail, send it.
    const code = findFreshOtp(config, state);
    if (code) {
      state.step = "DONE"; // mark as sent first -> the code is never sent twice
      cacheEngine.setProperty(item.key, JSON.stringify(state));
      sendSlackMessage(config.CHANNEL_ID, `<@${state.agent}>, here is your active OTP code: \`${code}\``, state.ts);
      continue;
    }

    if (age > MAX_WAIT_TIME_MS) {
      const mins = Math.round(MAX_WAIT_TIME_MS / 60000);
      state.step = "DONE";
      cacheEngine.setProperty(item.key, JSON.stringify(state));
      sendSlackMessage(config.CHANNEL_ID, `<@${state.agent}> No fresh OTP email arrived within ${mins} minutes. Please trigger 'token' again.`, state.ts);
    }
  }
}

// Looks in Gmail for a fresh 6-digit OTP for this request.
// Accepts emails that arrived shortly BEFORE the Slack request (agent connected first)
// or any time AFTER it. Each email is used only once.
// FIRST COME, FIRST SERVED: among all eligible emails the OLDEST one is returned,
// so the oldest waiting request gets the oldest OTP.
function findFreshOtp(config, state) {
  const props = PropertiesService.getScriptProperties();

  try {
    let usedIds = [];
    try { usedIds = JSON.parse(props.getProperty('used_otp_ids') || '[]'); } catch (e) { usedIds = []; }

    const slackRequestTimeMs = parseFloat(state.ts) * 1000;
    const emailFilters = config.EMAIL_SOURCES.map(src => `(from:${src} OR to:${src})`).join(" OR ");
    const searchQuery = `newer_than:1d (${emailFilters})`;

    const threads = GmailApp.search(searchQuery, 0, 5);
    let candidates = [];

    for (let thread of threads) {
      const messages = thread.getMessages();

      for (let i = messages.length - 1; i >= 0; i--) {
        const msg = messages[i];
        const msgId = msg.getId();
        if (usedIds.indexOf(msgId) !== -1) continue;

        const msgTimeMs = msg.getDate().getTime();
        const isNotTooOld = (msgTimeMs >= (slackRequestTimeMs - OTP_LOOKBACK_MS));
        const isFresh = (Date.now() - msgTimeMs) <= OTP_FRESH_MS;
        if (!isNotTooOld || !isFresh) continue;

        const fromText = msg.getFrom().toLowerCase();
        const isValidSource = config.EMAIL_SOURCES.some(src => fromText.includes(src.toLowerCase()));
        if (!isValidSource) continue;

        const match = msg.getSubject().match(/\b(\d{6})\b/) || msg.getPlainBody().match(/\b(\d{6})\b/);
        if (match) {
          candidates.push({ msg: msg, thread: thread, id: msgId, time: msgTimeMs, code: match[1] });
        }
      }
    }

    if (candidates.length === 0) return null;

    // Oldest eligible email first
    candidates.sort((a, b) => a.time - b.time);
    const pick = candidates[0];

    usedIds.push(pick.id);
    props.setProperty('used_otp_ids', JSON.stringify(usedIds.slice(-50)));
    pick.msg.markRead();
    pick.thread.markRead();
    return pick.code;
  } catch (e) {
    console.log("Gmail extraction error: " + e.toString());
  }
  return null;
}

function sendSlackMessage(channelId, messageText, targetThreadTs) {
  try {
    const postEndpoint = 'https://slack.com/api/chat.postMessage';
    const postPayload = { 'channel': channelId, 'text': String(messageText) };
    if (targetThreadTs) postPayload['thread_ts'] = String(targetThreadTs);
    
    const postOptions = {
      'method': 'post',
      'contentType': 'application/json; charset=utf-8',
      'headers': { 'Authorization': 'Bearer ' + SLACK_TOKEN },
      'payload': JSON.stringify(postPayload),
      'muteHttpExceptions': true
    };
    UrlFetchApp.fetch(postEndpoint, postOptions);
  } catch (e) { console.log('Slack sending error: ' + e.toString()); }
}

function getSlackUserName(userId) {
  const cacheEngine = PropertiesService.getScriptProperties();
  const cachedName = cacheEngine.getProperty(`user_name_${userId}`);
  if (cachedName) return cachedName;

  try {
    const userUrl = `https://slack.com/api/users.info?user=${userId}`;
    const httpOptions = {
      method: "get",
      headers: { "Authorization": "Bearer " + SLACK_TOKEN },
      muteHttpExceptions: true
    };
    
    const response = JSON.parse(UrlFetchApp.fetch(userUrl, httpOptions).getContentText());
    
    if (response.ok && response.user) {
      const fetchedName = response.user.real_name || response.user.profile.display_name || response.user.profile.real_name || response.user.name;
      if (fetchedName) {
        cacheEngine.setProperty(`user_name_${userId}`, fetchedName);
        return fetchedName;
      }
    } else {
      console.log(`Slack users.info Error for ${userId}: ` + JSON.stringify(response));
    }
  } catch (e) {
    console.log(`Failed to fetch user name for ${userId}: ` + e.toString());
  }

  return userId;
}

function getSlackChannelName(channelId) {
  const conf = BOT_CONFIGS.find(c => c.CHANNEL_ID === channelId);
  if (conf && conf.CHANNEL_NAME) return conf.CHANNEL_NAME;

  try {
    const url = `https://slack.com/api/conversations.info?channel=${channelId}`;
    const httpOptions = {
      method: "get",
      headers: { "Authorization": "Bearer " + SLACK_TOKEN },
      muteHttpExceptions: true
    };
    const res = JSON.parse(UrlFetchApp.fetch(url, httpOptions).getContentText());
    if (res.ok && res.channel && res.channel.name) {
      return "#" + res.channel.name;
    }
  } catch (e) {
    console.log(`Failed to fetch channel name for ${channelId}`);
  }
  return channelId;
}

// --- REAL-TIME WEAKSPOT DETECTOR & SECURITY ALERT SYSTEM ---

// Helper to send instant urgent alert when a security weakspot is found
function sendUrgentSecurityAlert(weakspotType, description, channelId, agentName) {
  const adminEmail = 'igeraldino@openaccessbpo.com';
  const timestamp = new Date().toLocaleString('en-US', { timeZone: 'Asia/Manila' });

  const alertSubject = `🚨 URGENT SECURITY WEAKSPOT ALERT: ${weakspotType}`;
  
  const alertHtml = `
  <div style="font-family: Arial, sans-serif; max-width: 600px; border: 2px solid #e53e3e; border-radius: 8px; overflow: hidden;">
    <div style="background-color: #e53e3e; color: white; padding: 12px 20px; font-weight: bold; font-size: 16px;">
      ⚠️ CRITICAL SECURITY WEAKSPOT DETECTED
    </div>
    <div style="padding: 20px; background-color: #fff5f5; color: #2d3748;">
      <p style="margin-top: 0;"><b>Detection Time:</b> ${timestamp}</p>
      <p><b>Threat Category:</b> <span style="color: #c53030; font-weight: bold;">${weakspotType}</span></p>
      <p><b>Affected Channel:</b> ${channelId}</p>
      <p><b>Agent / User Involved:</b> ${agentName || 'N/A'}</p>
      
      <div style="background: white; border-left: 4px solid #e53e3e; padding: 12px; margin: 15px 0; font-size: 13px;">
        <b>Technical Breakdown & Incident Detail:</b><br/>
        ${description}
      </div>

      <p style="font-size: 12px; color: #718096; margin-bottom: 0;">
        <b>Recommended Action:</b> Please log in to the Slack Admin Console / IT Admin Panel immediately to investigate and resolve this security anomaly.
      </p>
    </div>
  </div>
  `;

  // Send Email Alert
  GmailApp.sendEmail(adminEmail, alertSubject, "Critical Security Alert", { htmlBody: alertHtml });
  console.log(`[ALERT SENT] Security Weakspot Alert dispatched to ${adminEmail}`);
}


// --- AUTOMATED SHIFT REPORT VIA EMAIL (WITH COMPLETE ERROR AUDIT TRACKING) ---
function sendShiftReportEmail() {
  const recipientEmail = 'itdavao@openaccessbpo.com';
  const now = new Date();
 
  // ===================================================
  // 1. DETECT SHIFT (TRIGGER 1 HOUR BEFORE SHIFT OUT)
  // ===================================================
  const currentHour = now.getHours();
  let shiftName = "";
 
  // Morning Shift (6AM - 3PM) -> Triggered 1 hour before out (2:00 PM - 3:00 PM)
  if (currentHour >= 14 && currentHour < 15) {
    shiftName = "Morning Shift (6:00 AM - 3:00 PM)";
  }
  // Mid Shift (2PM - 11PM) -> Triggered 1 hour before out (10:00 PM - 11:00 PM)
  else if (currentHour >= 22 && currentHour < 23) {
    shiftName = "Mid Shift (2:00 PM - 11:00 PM)";
  }
  // Night Shift (10PM - 7AM) -> Triggered 1 hour before out (6:00 AM - 7:00 AM)
  else if (currentHour >= 6 && currentHour < 7) {
    shiftName = "Night Shift (10:00 PM - 7:00 AM)";
  }
  // IF MANUAL RUN / NOT IN SCHEDULED TRIGGER WINDOW -> CANCEL EXECUTION
  else {
    console.log(`[BLOCKED] Execution time (${now.toLocaleTimeString()}) is outside scheduled shift report windows (1 hr prior to shift out). Email skipped.`);
    return;
  }
 
  const nineHoursAgoSec = Math.floor((now.getTime() - (9 * 60 * 60 * 1000)) / 1000);
  const httpOptions = {
    method: "get",
    headers: { "Authorization": "Bearer " + SLACK_TOKEN },
    muteHttpExceptions: true
  };
 
  let grandTotalRequests = 0;
  let securityAnomaliesCount = 0;
  let detectedWeakspots = [];
  let channelData = [];
 
  // 2. COLLECT DATA FROM SLACK & AUDIT ALL ERRORS
  for (let config of BOT_CONFIGS) {
    const channelId = config.CHANNEL_ID;
    const channelName = getSlackChannelName(channelId);
    const historyUrl = `https://slack.com/api/conversations.history?channel=${channelId}&oldest=${nineHoursAgoSec}`;
 
    let channelRequests = 0;
    let channelLogs = [];
    let userRequestTimestamps = {};
 
    try {
      const fetchResult = UrlFetchApp.fetch(historyUrl, httpOptions);
      const response = JSON.parse(fetchResult.getContentText());
 
      if (!response.ok) {
        const errorMsg = `Slack API Error: ${response.error || 'Unknown'}`;
        securityAnomaliesCount++;
        detectedWeakspots.push({ channel: channelName, type: "API Auth Failure", detail: errorMsg });
        sendUrgentSecurityAlert("OAuth / API Failure", errorMsg, channelName, "System");
        continue;
      }
 
      if (response.messages) {
        const agentMessages = response.messages.filter(msg =>
          !msg.bot_id &&
          msg.text &&
          (msg.text.toLowerCase().includes('token') || msg.text.toLowerCase().includes('code'))
        );
 
        channelRequests = agentMessages.length;
        grandTotalRequests += channelRequests;
 
        agentMessages.sort((a, b) => parseFloat(a.ts) - parseFloat(b.ts));
 
        agentMessages.forEach(msg => {
          const msgTime = parseFloat(msg.ts);
          const timeStr = new Date(msgTime * 1000).toLocaleTimeString('en-US', {
            hour: '2-digit',
            minute: '2-digit',
            hour12: true,
            timeZone: 'Asia/Manila'
          });
 
          const agentName = getSlackUserName(msg.user);
 
          // WEAKSPOT: RATE LIMITING (3+ requests within 5 mins)
          if (!userRequestTimestamps[msg.user]) userRequestTimestamps[msg.user] = [];
          userRequestTimestamps[msg.user].push(msgTime);
 
          const recentRequests = userRequestTimestamps[msg.user].filter(t => (msgTime - t) <= 300);
          if (recentRequests.length > 3) {
            securityAnomaliesCount++;
            const floodDetail = `User <b>${agentName}</b> submitted ${recentRequests.length} OTP requests within 5 minutes.`;
            detectedWeakspots.push({ channel: channelName, type: "Excessive Token Requests", detail: floodDetail });
            sendUrgentSecurityAlert("Excessive Request Flood", floodDetail, channelName, agentName);
          }
 
          // WEAKSPOT: UNKNOWN / UNRESOLVED USER
          // FIXED: getSlackUserName() never returns a string starting with
          // "Unknown User" — its only fallback is the raw userId. The old
          // `.startsWith("Unknown User")` check was dead code and never
          // fired. The real fallback signal is `agentName === msg.user`.
          if (agentName === msg.user) {
            securityAnomaliesCount++;
            const unkDetail = `Unrecognized user ID: <code>${msg.user}</code>.`;
            detectedWeakspots.push({ channel: channelName, type: "Unverified User Request", detail: unkDetail });
          }
 
          let itResponder = "Auto / Bot";
          let threadStatus = "Fulfilled";
 
          // CHECK THREAD REPLIES & SCAN FOR TIMEOUT/ERRORS
          if (msg.reply_count && msg.reply_count > 0) {
            try {
              const repliesUrl = `https://slack.com/api/conversations.replies?channel=${channelId}&ts=${msg.ts}`;
              const repliesRes = JSON.parse(UrlFetchApp.fetch(repliesUrl, httpOptions).getContentText());
 
              if (repliesRes.ok && repliesRes.messages) {
                for (let reply of repliesRes.messages) {
                  const replyText = reply.text || "";
 
                  // DETECT TIMEOUT ERROR
                  if (replyText.includes("timed out") || replyText.includes("didn't see you type")) {
                    threadStatus = "Error: Request Timed Out";
                    securityAnomaliesCount++;
                    detectedWeakspots.push({
                      channel: channelName,
                      type: "Process Timeout Error",
                      detail: `Agent <b>${agentName}</b> request timed out (No 'sent'/'done' response within ${Math.round(MAX_WAIT_TIME_MS / 60000)} mins).`
                    });
                  }
                  // DETECT NO OTP ARRIVED ERROR
                  else if (replyText.includes("No fresh OTP email arrived")) {
                    threadStatus = "Error: OTP Email Missing";
                    securityAnomaliesCount++;
                    detectedWeakspots.push({
                      channel: channelName,
                      type: "Gmail Extraction Failed",
                      detail: `No OTP email found for agent <b>${agentName}</b> within the expected window.`
                    });
                  }
 
                  if (reply.ts === msg.ts || reply.bot_id || reply.subtype === "bot_message") continue;
 
                  const responderName = getSlackUserName(reply.user);
                  if (/\bIT\b/i.test(responderName) || responderName.toLowerCase().includes("it ")) {
                    itResponder = responderName;
                    break;
                  } else if (reply.user !== msg.user) {
                    itResponder = responderName;
                  }
                }
              }
            } catch (e) {
              console.log(`Failed to fetch replies for thread ${msg.ts}`);
              securityAnomaliesCount++;
              detectedWeakspots.push({
                channel: channelName,
                type: "Thread Fetch Exception",
                detail: `Error reading thread ${msg.ts}: ${e.toString()}`
              });
            }
          } else {
            // UNRESOLVED / NO BOT RESPONSE
            threadStatus = "Pending / Unfulfilled";
          }
 
          channelLogs.push({
            time: timeStr,
            agent: agentName,
            responder: itResponder,
            status: threadStatus
          });
        });
      }
    } catch (e) {
      securityAnomaliesCount++;
      const sysErr = `Exception caught: ${e.toString()}`;
      detectedWeakspots.push({ channel: channelName, type: "System Exception", detail: sysErr });
      sendUrgentSecurityAlert("System Exception Error", sysErr, channelName, "System");
    }
 
    channelData.push({
      channelId: channelId,
      channelName: channelName,
      total: channelRequests,
      logs: channelLogs
    });
  }
 
  // 3. GENERATE WEAKSPOT & ERROR SUMMARY HTML
  let weakspotsSummaryHtml = '';
  if (detectedWeakspots.length > 0) {
    let listItems = detectedWeakspots.map(w => `
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #fecdd3; vertical-align: top;">
          <span style="display: inline-block; background: linear-gradient(135deg, #f43f5e, #dc2626); color: #ffffff; font-size: 9px; font-weight: 700; letter-spacing: 0.4px; padding: 3px 8px; border-radius: 10px; white-space: nowrap;">${w.type.toUpperCase()}</span>
        </td>
        <td style="padding: 8px 0 8px 12px; border-bottom: 1px solid #fecdd3; font-size: 12.5px; color: #44403c;">
          <b style="color: #18181b;">${w.channel}</b> — ${w.detail}
        </td>
      </tr>
    `).join('');
 
    weakspotsSummaryHtml = `
      <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; margin-bottom: 20px; background: linear-gradient(180deg, #fff1f2, #fef2f2); border: 1px solid #fecdd3; border-radius: 14px;">
        <tr>
          <td style="padding: 14px 18px 4px 18px;">
            <span style="font-size: 13.5px; font-weight: 800; color: #be123c;">⚠️ ${detectedWeakspots.length} System Error${detectedWeakspots.length === 1 ? '' : 's'} / Security Flag${detectedWeakspots.length === 1 ? '' : 's'}</span>
          </td>
        </tr>
        <tr>
          <td style="padding: 2px 18px 14px 18px;">
            <table border="0" cellpadding="0" cellspacing="0" style="width: 100%;">
              ${listItems}
            </table>
          </td>
        </tr>
      </table>
    `;
  } else {
    weakspotsSummaryHtml = `
      <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; margin-bottom: 20px; background: linear-gradient(180deg, #f0fdf4, #ecfdf5); border: 1px solid #a7f3d0; border-radius: 14px;">
        <tr>
          <td style="padding: 16px 18px; font-size: 12.5px; color: #166534;">
            <span style="font-size: 14px; font-weight: 800;">✅ No Errors or Security Weakspots Detected</span><br/>
            <span style="color: #15803d;">All token operations and request rates complied fully with IT Security Policies this shift.</span>
          </td>
        </tr>
      </table>
    `;
  }
 
  // 4. GENERATE VISUAL ANALYTICS GRAPH
  let analyticsRowsHtml = '';
  channelData.forEach((item, idx) => {
    const percentage = grandTotalRequests > 0 ? Math.round((item.total / grandTotalRequests) * 100) : 0;
    const barWidth = percentage > 0 ? percentage : 2;
    const rowBg = idx % 2 === 0 ? '#ffffff' : '#f8fafc';
 
    analyticsRowsHtml += `
      <tr style="background-color: ${rowBg};">
        <td style="padding: 12px 14px; font-weight: 700; color: #1e293b; font-size: 12.5px; width: 38%;">${item.channelName}<br/><span style="font-size: 10px; color: #94a3b8; font-weight: normal;">${item.channelId}</span></td>
        <td style="padding: 12px 14px; width: 12%; text-align: center;">
          <span style="display: inline-block; background: linear-gradient(135deg, #e0e7ff, #c7d2fe); color: #3730a3; font-size: 12px; font-weight: 800; padding: 4px 10px; border-radius: 10px;">${item.total}</span>
        </td>
        <td style="padding: 12px 14px; width: 50%;">
          <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; background-color: #e2e8f0; border-radius: 8px;">
            <tr>
              <td style="width: ${barWidth}%; background: linear-gradient(90deg, #4f46e5, #2563eb); height: 18px; border-radius: 8px; text-align: right; color: #ffffff; font-size: 9px; font-weight: 800; padding-right: 7px; line-height: 18px;">
                ${percentage > 10 ? percentage + '%' : ''}
              </td>
              <td style="width: ${100 - barWidth}%;"></td>
            </tr>
          </table>
        </td>
      </tr>`;
  });
 
  // 5. GENERATE DETAILED CHANNEL LOG TABLES
  const STATUS_STYLES = {
    fulfilled: { bg: '#dcfce7', fg: '#166534' },
    pending:   { bg: '#fef3c7', fg: '#92400e' },
    error:     { bg: '#fee2e2', fg: '#b91c1c' }
  };
  function statusBadge(status) {
    let style = STATUS_STYLES.fulfilled;
    if (status.startsWith("Error")) style = STATUS_STYLES.error;
    else if (status.includes("Pending")) style = STATUS_STYLES.pending;
    return `<span style="display: inline-block; background-color: ${style.bg}; color: ${style.fg}; font-size: 10px; font-weight: 700; padding: 3px 9px; border-radius: 10px; white-space: nowrap;">${status}</span>`;
  }
 
  let channelLogsHtml = '';
  channelData.forEach(item => {
    let rowsHtml = '';
    if (item.logs.length > 0) {
      item.logs.forEach((log, index) => {
        const bg = index % 2 === 0 ? '#ffffff' : '#f8fafc';
        const isBot = log.responder === "Auto / Bot";
        const responderDisplay = isBot
          ? `<span style="color: #94a3b8; font-style: italic; font-size: 11px;">Auto / Bot</span>`
          : `<b style="color: #2563eb; font-size: 11px;">${log.responder}</b>`;
 
        rowsHtml += `
          <tr style="background-color: ${bg};">
            <td style="padding: 9px 12px; border-bottom: 1px solid #f1f5f9; font-size: 12px; color: #64748b; width: 18%;">${log.time}</td>
            <td style="padding: 9px 12px; border-bottom: 1px solid #f1f5f9; font-size: 12px; font-weight: 600; color: #1e293b; width: 32%;">${log.agent}</td>
            <td style="padding: 9px 12px; border-bottom: 1px solid #f1f5f9; font-size: 12px; width: 25%;">${responderDisplay}</td>
            <td style="padding: 9px 12px; border-bottom: 1px solid #f1f5f9; text-align: right; width: 25%;">${statusBadge(log.status)}</td>
          </tr>`;
      });
    } else {
      rowsHtml = `
        <tr>
          <td colspan="4" style="padding: 18px; text-align: center; color: #94a3b8; font-style: italic; font-size: 12px;">
            No token / OTP requests recorded in this channel during this shift.
          </td>
        </tr>`;
    }
 
    channelLogsHtml += `
      <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; margin-bottom: 18px; border: 1px solid #e2e8f0; border-radius: 12px; border-collapse: separate; overflow: hidden;">
        <thead>
          <tr style="background: linear-gradient(90deg, #0f172a, #1e293b);">
            <th colspan="3" style="padding: 11px 16px; text-align: left; font-size: 13px; font-weight: 700; color: #ffffff;">${item.channelName} <span style="font-weight: 400; color: #94a3b8; font-size: 11px;">(${item.channelId})</span></th>
            <th style="padding: 11px 16px; text-align: right; font-size: 12px; color: #ffffff;">
              <span style="background-color: #334155; padding: 3px 9px; border-radius: 10px;">Total ${item.total}</span>
            </th>
          </tr>
          <tr style="background-color: #f8fafc; text-align: left; font-size: 10px; color: #64748b; text-transform: uppercase; letter-spacing: 0.4px;">
            <th style="padding: 8px 12px;">Timestamp</th>
            <th style="padding: 8px 12px;">Agent Name</th>
            <th style="padding: 8px 12px;">IT Responder</th>
            <th style="padding: 8px 12px; text-align: right;">Status</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
        </tbody>
      </table>`;
  });
 
  // 6. ASSEMBLE HTML TEMPLATE  (UI-only rework — logic above is untouched)
  const overallStatusBadge = securityAnomaliesCount === 0
    ? `<span style="display: inline-block; background: linear-gradient(135deg, #22c55e, #16a34a); color: #ffffff; font-size: 12px; font-weight: 800; padding: 5px 14px; border-radius: 14px;">🟢 HEALTHY — 100% Success Rate</span>`
    : `<span style="display: inline-block; background: linear-gradient(135deg, #f87171, #dc2626); color: #ffffff; font-size: 12px; font-weight: 800; padding: 5px 14px; border-radius: 14px;">🔴 NEEDS ATTENTION — ${securityAnomaliesCount} Issue${securityAnomaliesCount === 1 ? '' : 's'}</span>`;
 
  // Friendly little sign-off line from the "bot", tone changes with shift health
  const botMascotLine = securityAnomaliesCount === 0
    ? "Beep boop — smooth shift, nothing to report! 🤖"
    : "Beep boop — I found a few things that need a human look. 🤖";
 
  const htmlTemplate = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
  </head>
  <body style="font-family: 'Segoe UI', Helvetica, Arial, sans-serif; background-color: #eef1f6; margin: 0; padding: 18px; color: #1e293b;">
    <table border="0" cellpadding="0" cellspacing="0" style="max-width: 700px; width: 100%; margin: 0 auto; background-color: #ffffff; border-radius: 18px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 10px 30px rgba(15,23,42,0.08);">
      <tr>
        <td>
 
          <!-- HEADER -->
          <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; background: linear-gradient(120deg, #0f172a 0%, #1e293b 60%, #1d3a6b 100%);">
            <tr>
              <td style="padding: 24px 26px 20px 26px;">
                <table border="0" cellpadding="0" cellspacing="0" style="width: 100%;">
                  <tr>
                    <td style="vertical-align: middle; width: 104px;">
                      <img src="cid:oabLogo" width="96" height="48" alt="Open Access BPO" style="display: block; border-radius: 8px; background-color: #ffffff; padding: 6px 10px;" />
                    </td>
                    <td style="vertical-align: middle; padding-left: 12px;">
                      <span style="display: block; font-size: 10px; font-weight: 700; letter-spacing: 1.2px; color: #60a5fa; text-transform: uppercase; margin-bottom: 3px;">IT Davao • OTP Bot Engine</span>
                      <span style="display: block; font-size: 19px; font-weight: 800; color: #ffffff;">Shift Handover &amp; System Audit Report</span>
                    </td>
                    <td style="vertical-align: middle; text-align: right; white-space: nowrap; width: 158px;">
                      <img src="cid:fortinetLogo" width="150" height="20" alt="Fortinet" style="display: block; margin-left: auto;" />
                    </td>
                  </tr>
                </table>
                <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; margin-top: 12px;">
                  <tr>
                    <td>
                      <span style="display: inline-block; background-color: #1e293b; color: #cbd5e1; font-size: 11px; padding: 5px 12px; border-radius: 12px; border: 1px solid #334155;">${shiftName}</span>
                      <span style="display: inline-block; color: #94a3b8; font-size: 11px; margin-left: 10px;">${now.toLocaleDateString('en-US', { timeZone: 'Asia/Manila', weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</span>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
 
          <table border="0" cellpadding="0" cellspacing="0" style="width: 100%;">
            <tr>
              <td style="padding: 24px 22px 4px 22px;">
 
                <!-- METRIC CARDS -->
                <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; margin-bottom: 20px;">
                  <tr>
                    <td style="width: 4px;"></td>
                    <td style="width: 32%; background: linear-gradient(160deg, #eff6ff, #dbeafe); border: 1px solid #bfdbfe; border-radius: 12px; padding: 16px; text-align: center;">
                      <span style="font-size: 10px; color: #2563eb; font-weight: 800; letter-spacing: 0.4px; display: block; text-transform: uppercase;">Total Requests</span>
                      <span style="font-size: 28px; font-weight: 800; color: #1d4ed8; display: block; margin-top: 4px;">${grandTotalRequests}</span>
                    </td>
                    <td style="width: 4px;"></td>
                    <td style="width: 32%; background: linear-gradient(160deg, #f0fdf4, #dcfce7); border: 1px solid #bbf7d0; border-radius: 12px; padding: 16px; text-align: center;">
                      <span style="font-size: 10px; color: #16a34a; font-weight: 800; letter-spacing: 0.4px; display: block; text-transform: uppercase;">Monitored Channels</span>
                      <span style="font-size: 28px; font-weight: 800; color: #15803d; display: block; margin-top: 4px;">${BOT_CONFIGS.length}</span>
                    </td>
                    <td style="width: 4px;"></td>
                    <td style="width: 32%; background: ${securityAnomaliesCount > 0 ? 'linear-gradient(160deg, #fef2f2, #fee2e2)' : 'linear-gradient(160deg, #faf5ff, #f3e8ff)'}; border: 1px solid ${securityAnomaliesCount > 0 ? '#fecaca' : '#e9d5ff'}; border-radius: 12px; padding: 16px; text-align: center;">
                      <span style="font-size: 10px; color: ${securityAnomaliesCount > 0 ? '#dc2626' : '#7c3aed'}; font-weight: 800; letter-spacing: 0.4px; display: block; text-transform: uppercase;">System Health</span>
                      <span style="font-size: 15px; font-weight: 800; color: ${securityAnomaliesCount > 0 ? '#b91c1c' : '#6d28d9'}; display: block; margin-top: 6px;">${securityAnomaliesCount > 0 ? `${securityAnomaliesCount} Flagged` : 'All Normal'}</span>
                    </td>
                    <td style="width: 4px;"></td>
                  </tr>
                </table>
 
                <!-- WEAKSPOT & ERROR SUMMARY SECTION -->
                ${weakspotsSummaryHtml}
 
                <!-- VISUAL GRAPH -->
                <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; margin-bottom: 22px; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">
                  <thead>
                    <tr>
                      <th colspan="3" style="background-color: #f8fafc; padding: 11px 14px; text-align: left; font-size: 12px; font-weight: 800; color: #1e293b; border-bottom: 1px solid #e2e8f0;">📈 Channel Activity Distribution</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${analyticsRowsHtml}
                  </tbody>
                </table>
 
                <!-- DETAILED LOGS -->
                <h3 style="margin: 0 0 12px 0; font-size: 14px; color: #1e293b; font-weight: 800;">Detailed Activity &amp; Audit Logs per Campaign</h3>
                ${channelLogsHtml}
 
                <!-- OVERALL REVIEW & RECOMMENDATION SECTION -->
                <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; margin: 8px 0 20px 0; background: linear-gradient(160deg, #f8fafc, #f1f5f9); border: 1px solid #e2e8f0; border-radius: 12px;">
                  <tr>
                    <td style="padding: 18px 20px;">
                      <span style="font-size: 12px; font-weight: 800; color: #1e293b; text-transform: uppercase; letter-spacing: 0.4px;">📊 Overall Operational Review</span>
 
                      <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; font-size: 12px; color: #334155; margin-top: 12px;">
                        <tr>
                          <td style="padding-bottom: 10px;">${overallStatusBadge}</td>
                        </tr>
                        <tr>
                          <td style="padding-bottom: 10px; line-height: 1.55;">
                            During this shift, the system processed a total of <b>${grandTotalRequests} OTP/Token requests</b> across <b>${BOT_CONFIGS.length} monitored channels</b>.
                            ${securityAnomaliesCount > 0
                              ? `There were <b style="color: #dc2626;">${securityAnomaliesCount} system flags/errors</b> recorded (timeouts, missing Gmail OTPs, or API throttles).`
                              : 'All agent requests were successfully matched and fulfilled without critical delays or system errors.'
                            }
                          </td>
                        </tr>
                        <tr>
                          <td>
                            <b>Recommended Action Items:</b>
                            <ul style="margin: 6px 0 0 0; padding-left: 18px; color: #475569; line-height: 1.6;">
                              ${securityAnomaliesCount > 0
                                ? `
                                  <li>Review the flagged entries above and verify the relevant 2FA email sources are active.</li>
                                  <li>If "Gmail Extraction Failed" recurs, confirm whether the mail server is delaying OTP delivery beyond the freshness window.</li>
                                `
                                : `
                                  <li>Maintain current setup — bot operations are running optimally. No intervention required for this shift.</li>
                                `
                              }
                              <li>Remind agents to avoid multiple 'token' triggers within 5 minutes to prevent rate-limit flags.</li>
                            </ul>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
 
              </td>
            </tr>
          </table>
 
          <!-- FOOTER -->
          <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; background-color: #f8fafc; border-top: 1px solid #e2e8f0;">
            <tr>
              <td style="padding: 16px 22px; text-align: center;">
                <span style="display: block; color: #64748b; font-size: 11.5px; font-style: italic; margin-bottom: 6px;">${botMascotLine}</span>
                <span style="color: #94a3b8; font-size: 11px;">Automated Report generated by <b style="color: #64748b;">IT Davao OTP Bot Engine</b> &middot; Open Access BPO</span>
              </td>
            </tr>
          </table>
 
        </td>
      </tr>
    </table>
  </body>
  </html>`;
 
  // 7. SEND EMAIL
  const emailSubject = `[SLACK BOT & SYSTEM AUDIT] Multi-Channel Activity Report - ${shiftName}`;
 
  // Open Access BPO logo, embedded inline (no external hosting needed) and
  // referenced above via src="cid:oabLogo".
  const oabLogoBlob = Utilities.newBlob(
    Utilities.base64Decode(OAB_LOGO_BASE64),
    'image/png',
    'oab-logo.png'
  ).setName('oabLogo');
 
  const fortinetLogoBlob = Utilities.newBlob(
    Utilities.base64Decode(FORTINET_LOGO_BASE64),
    'image/png',
    'fortinet-logo.png'
  ).setName('fortinetLogo');

  GmailApp.sendEmail(recipientEmail, emailSubject, "Your email client does not support HTML emails.", {
    htmlBody: htmlTemplate,
    inlineImages: { oabLogo: oabLogoBlob, fortinetLogo: fortinetLogoBlob }
  });
 
  console.log(`Detailed error-audit report successfully sent to ${recipientEmail} for ${shiftName}`);
}

// ===================================================
// Open Access BPO logo, base64-encoded PNG.
// Kept as a top-level constant so sendShiftReportEmail() stays readable.
// Swap this out any time by re-exporting a new logo as base64.
// ===================================================
const OAB_LOGO_BASE64 = "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";

// Fortinet logo (white wordmark, transparent PNG), shown at the top-right of the report header.
const FORTINET_LOGO_BASE64 = "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";

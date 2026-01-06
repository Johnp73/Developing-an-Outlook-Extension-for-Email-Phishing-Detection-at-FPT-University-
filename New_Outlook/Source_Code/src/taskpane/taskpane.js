/* global Office, document */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Đăng ký trình xử lý cho ItemChanged
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, handleItemChanged);

    // Kiểm tra email hiện tại khi tải
    checkCurrentItem();

    // Thêm trình lắng nghe sự kiện cho nút hiển thị tệp ẩn
    document.getElementById("show-hidden-btn").addEventListener("click", handleShowHidden);

    // Thêm trình lắng nghe sự kiện cho nút kiểm tra AI
    document.getElementById("run-ai-check-btn").addEventListener("click", handleAiCheck);

    // Khởi tạo Pivot (các tab)
    initPivot();
  }
});

let isShowingHidden = false;
let isRunningAiCheck = false; 

function initPivot() {
  const pivotLinks = document.querySelectorAll(".ms-Pivot-link");
  pivotLinks.forEach(link => {
    link.addEventListener("click", (e) => {
      pivotLinks.forEach(l => l.classList.remove("is-selected"));
      e.target.classList.add("is-selected");

      const contents = document.querySelectorAll(".ms-Pivot-content");
      contents.forEach(c => c.style.display = "none");

      const selectedContent = document.querySelector(`.ms-Pivot-content[data-content="${e.target.dataset.content}"]`);
      if (selectedContent) {
        selectedContent.style.display = "block";
      }
    });
  });

  // Hiển thị tab đầu tiên mặc định
  pivotLinks[0].click();
}

function handleItemChanged(eventArgs) {
  // Kiểm tra email mới khi item thay đổi
  checkCurrentItem();

  // Tự động cập nhật kết quả nếu đang ở chế độ hiển thị tệp ẩn
  if (isShowingHidden) {
    updateHiddenResults();
  }

  // Tự động cập nhật kiểm tra AI nếu đang ở chế độ AI
  if (isRunningAiCheck) {
    runAiPhishingCheck();
  }
}

async function checkCurrentItem() {
  const item = Office.context.mailbox.item;
  if (item && item.itemType === Office.MailboxEnums.ItemType.Message) {
    const senderEmail = item.sender.emailAddress.toLowerCase();
    const domain = senderEmail.split('@')[1];

    // Sử dụng các domain được mã hóa cứng (tải trực tiếp từ nội dung JSON)
    let whitelistDomains = domains.whitelist || [];
    let blacklistDomains = domains.blacklist || [];

    // Xóa các thông báo cũ (thêm key mới để xóa)
    await new Promise((resolve) => {
      item.notificationMessages.removeAsync("PhishingWarning", resolve);
    });
    await new Promise((resolve) => {
      item.notificationMessages.removeAsync("SafeNotification", resolve);
    });
    await new Promise((resolve) => {
      item.notificationMessages.removeAsync("UnknownNotification", resolve);
    });
    await new Promise((resolve) => {
      item.notificationMessages.removeAsync("LinkWarning", resolve); // Xóa thông báo link riêng nếu có
    });

    let messageType, messageText, icon, key, persistent;

    if (whitelistDomains.includes(domain)) {
      messageType = Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;
      messageText = `Email thuộc tổ chức đại học FPT.`;
      icon = "Icon.80x80"; // Hoặc sử dụng icon an toàn như 'ms-Icon--Completed'
      key = "SafeNotification";
      persistent = false;
    } else if (blacklistDomains.includes(domain)) {
      messageType = Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage; // Hoặc sử dụng ErrorMessage để cảnh báo mạnh hơn
      messageText = `Email lừa đảo, nguy hiểm.`;
      icon = "Icon.80x80"; // Hoặc sử dụng icon cảnh báo như 'ms-Icon--Warning'
      key = "PhishingWarning";
      persistent = true;
    } else {
      messageType = Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;
      messageText = `Thận trọng: Người gửi không có trong danh sách.`;
      icon = "Icon.80x80"; // Hoặc sử dụng icon thông tin như 'ms-Icon--Info'
      key = "UnknownNotification";
      persistent = false;
    }

    // Lấy nội dung email để kiểm tra link
    let hasLinks = false;
    let hasDangerousLinks = false;
    const bodyResult = await new Promise((resolve) => {
      item.body.getAsync("html", resolve);
    });
    if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
      const bodyHtml = bodyResult.value;
      const urls = extractUrlsFromHtml(bodyHtml);
      hasLinks = urls.length > 0;
      if (hasLinks) {
        for (const url of urls) {
          const pred = await predictUrl(url, bodyHtml);
          if (pred.status === 'Nguy hiểm' || pred.status === 'Lỗi dự đoán') {
            hasDangerousLinks = true;
            break; // Dừng sớm nếu tìm thấy bất kỳ link nguy hiểm nào
          }
        }
      }
    }

    // Cắt ngắn thông báo chính nếu cần (dù ít xảy ra vì không nối thêm nữa)
    if (messageText.length > 150) {
      messageText = messageText.substring(0, 147) + '...';
    }

    // Thêm thông báo chính cho domain
    const mainMessage = {
      type: messageType,
      message: messageText,
      icon: icon,
      persistent: persistent,
    };
    await new Promise((resolve) => {
      item.notificationMessages.replaceAsync(key, mainMessage, resolve);
    });

    // Nếu có link, thêm thông báo riêng dựa trên kiểm tra AI
    if (hasLinks) {
      let linkMessageText;
      if (hasDangerousLinks) {
        linkMessageText = "Cảnh báo: Email chứa link nguy hiểm";
      } else {
        linkMessageText = "Email chứa link an toàn";
      }
      // Cắt ngắn nếu cần
      if (linkMessageText.length > 150) {
        linkMessageText = linkMessageText.substring(0, 147) + '...';
      }
      const linkMessage = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, // Hoặc ErrorMessage nếu muốn cảnh báo mạnh hơn
        message: linkMessageText,
        icon: "Icon.80x80", // Có thể dùng icon cảnh báo như 'ms-Icon--Warning'
        persistent: true, // Giữ hiển thị lâu
      };
      await new Promise((resolve) => {
        item.notificationMessages.replaceAsync("LinkWarning", linkMessage, resolve);
      });
    }
  }
}

function splitConcatenatedUrls(str) {
  return str.split(/(?=https?:\/\/)/g).filter(u => u.trim() && u.startsWith('http'));
}

function cleanUrl(url) {
  // Cắt bỏ khoảng trắng đầu/cuối
  url = url.trim();

  // Thay thế các thực thể HTML như &amp; thành & (thường gặp trong HTML email)
  url = url.replace(/&amp;/gi, '&');

  // Xóa dấu chấm câu cuối nếu không phải phần của URL (ví dụ: dấu chấm, dấu phẩy ở cuối)
  url = url.replace(/[\.,;:!?]*$/, '');

  // Giải mã các ký tự được mã hóa URL (ví dụ: %3F thành ?)
  try {
    url = decodeURI(url);
  } catch (e) {
    // Nếu giải mã thất bại, giữ nguyên bản gốc
  }

  // Chuẩn hóa giao thức và tên máy chủ thành chữ thường
  try {
    const urlObj = new URL(url);
    urlObj.protocol = urlObj.protocol.toLowerCase();
    urlObj.hostname = urlObj.hostname.toLowerCase();

    // Xóa cổng mặc định (80 cho http, 443 cho https)
    if ((urlObj.protocol === 'http:' && urlObj.port === '80') ||
        (urlObj.protocol === 'https:' && urlObj.port === '443')) {
      urlObj.port = '';
    }

    // Trả về href đã chuẩn hóa mà không có dấu gạch chéo cuối nếu không cần
    url = urlObj.href.replace(/\/$/, '');
  } catch (e) {
    // Nếu không phải URL hợp lệ, trả về bản gốc đã cắt
    return url;
  }

  return url;
}

// Hàm cải tiến để trích xuất URL từ HTML (trích xuất href từ <a> và regex tốt hơn cho văn bản thuần)
function extractUrlsFromHtml(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  let links = Array.from(doc.querySelectorAll('a[href]')).map(a => a.href);

  // Áp dụng split và clean cho các link, phòng trường hợp
  links = links.flatMap(splitConcatenatedUrls).map(cleanUrl);

  // Trích xuất URL thuần từ văn bản sử dụng regex trên toàn bộ văn bản
  const text = doc.body.innerText;
  const urlRegex = /(https?:\/\/(?:www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}[-a-zA-Z0-9()@:%_\+.~#?&\/=]*)/gi;
  const plainMatches = text.match(urlRegex) || [];
  const plainUrls = plainMatches.flatMap(splitConcatenatedUrls).map(cleanUrl);

  // Kết hợp và loại bỏ trùng lặp, với xác thực
  const allUrlsSet = new Set([...links, ...plainUrls]);
  const allUrls = Array.from(allUrlsSet).filter(url => {
    try {
      new URL(url); 
      return url.startsWith('http') && url.length > 10 && !url.includes(' ');
    } catch {
      return false;
    }
  });

  return allUrls;
}


function extractFeatures(url, bodyHtml) {
  let features = new Array(25).fill(0); 
  try {
    const urlObj = new URL(url);
    const hostname = urlObj.hostname;
    const domain = hostname; // Để đơn giản

    let text = bodyHtml.toLowerCase();

    // 0: IP_Address (1 nếu IP, -1 nếu không)
    const ip = /^((\d{1,3}\.){3}\d{1,3})$/.test(domain) || /^([0-9a-fA-F:]+)$/.test(domain); // IPv4 hoặc IPv6
    features[0] = ip ? 1 : -1;

    // 1: URL_Length (-1 nếu <54, 0 nếu 54-75, 1 nếu >75)
    const urlLen = url.length;
    features[1] = urlLen < 54 ? -1 : (urlLen <= 75 ? 0 : 1);

    // 2: Shortining_Service (1 nếu khớp mẫu, -1 nếu không)
    const shorteningPattern = /bit\.ly|goo\.gl|shorte\.st|go2l\.ink|x\.co|ow\.ly|t\.co|tinyurl|tr\.im|is\.gd|cli\.gs/i;
    features[2] = shorteningPattern.test(url) ? 1 : -1;

    // 3: having_At_Symbol (1 nếu có @, -1 nếu không)
    features[3] = url.includes('@') ? 1 : -1;

    // 4: double_slash_redirecting (-1 nếu // cuối cùng <7, 1 nếu không)
    const lastDoubleSlash = url.lastIndexOf('//');
    features[4] = lastDoubleSlash < 7 ? -1 : 1;

    // 5: Prefix_Suffix (1 nếu có - trong domain, -1 nếu không)
    features[5] = domain.includes('-') ? 1 : -1;

    // 6: having_Sub_Domain (-1 nếu 1 dấu chấm, 0 nếu 2, 1 nếu >2)
    const dots = domain.split('.').length - 1;
    features[6] = dots === 1 ? -1 : (dots === 2 ? 0 : 1);

    // 7: SSLfinal_State (-1 nếu https, 1 nếu không)
    features[7] = urlObj.protocol === 'https:' ? -1 : 1;

    // 8: Domain_registeration_length 
    features[8] = 0; 

    // 9: Favicon 
    features[9] = 0;

    // 10: port (1 nếu cổng không chuẩn, -1 nếu không)
    const port = urlObj.port;
    features[10] = port && port !== '80' && port !== '443' ? 1 : -1;

    // 11: HTTPS_token (1 nếu 'https' trong domain, -1 nếu không)
    features[11] = domain.toLowerCase().includes('https') ? 1 : -1;

    // 12: Request_URL 
    features[12] = 0;

    // 13: URL_of_Anchor 
    let totalAnchors = (text.match(/https?:\/\//gi) || []).length;
    let unsafeAnchors = (text.match(/(#|javascript:|about:blank)/gi) || []).length;
    const domainMatches = text.match(/https?:\/\/([^\s\/]+)/gi) || [];
    let externalCount = 0;
    for (const match of domainMatches) {
      let group1 = match.replace(/https?:\/\//i, '');
      if (!group1.includes(domain)) {
        externalCount++;
      }
    }
    unsafeAnchors += externalCount;
    let percent = totalAnchors > 0 ? (unsafeAnchors / totalAnchors) * 100 : 100;
    features[13] = percent < 31 ? -1 : (percent < 67 ? 0 : 1);

    // 14: Links_in_tags 
    let totalTags = (text.match(/<meta|<script|<link/gi) || []).length;
    let external = (text.match(/https?:\/\/([^"']+)/gi) || []).length;
    percent = totalTags > 0 ? (external / totalTags) * 100 : 100;
    features[14] = percent < 17 ? -1 : (percent < 81 ? 0 : 1);

    // 15: SFH 
    const emptyActionRegex = /<form[^>]*action\s*=\s*["']\s*["']/i;
    if (emptyActionRegex.test(text)) {
      features[15] = 1;
    } else {
      const formActionRegex = /<form[^>]*action=["']https?:\/\/([^\/'"]+)["']/i;
      const match = formActionRegex.exec(text);
      if (match) {
        const formDomain = match[1];
        features[15] = formDomain.includes(domain) ? -1 : 0;
      } else {
        features[15] = -1;
      }
    }

    // 16: Submitting_to_email 
    features[16] = text.includes("mailto:") || text.includes("mail()") ? 1 : -1;

    // 17: Abnormal_URL 
    features[17] = 0;

    // 18: Redirect 
    features[18] = 0;

    // 19: on_mouseover 
    features[19] = /onmouseover/gi.test(text) ? 1 : -1;

    // 20: RightClick 
    features[20] = /event\.button\s*==\s*2/gi.test(text) ? 1 : -1;

    // 21: popUpWidnow 
    features[21] = text.includes("alert(") ? 1 : -1;

    // 22: Iframe 
    features[22] = /<iframe/gi.test(text) ? 1 : -1;

    // 23: age_of_domain 
    features[23] = 0; 

    // 24: DNSRecord 
    features[24] = 0; 

    return features;
  } catch (error) {
    console.error('Lỗi trích xuất đặc trưng:', error);
    // Trả về đặc trưng lừa đảo mặc định khi lỗi
    return new Array(25).fill(1);
  }
}

// Hàm đã sửa để dự đoán xem URL có phải lừa đảo không (cục bộ sử dụng rf_model.js)
async function predictUrl(url, bodyHtml) {
  try {
    const input = extractFeatures(url, bodyHtml); // Từ rf_model.js, trả về [prob_legitimate, prob_phishing]
    const modelScores = score(input);
    console.log(modelScores);

    // Giải thích: Nếu prob_phishing
    return modelScores[1] > modelScores[0]
      ? { status: 'Nguy hiểm', icon: 'ms-Icon--Cancel' } 
      : { status: 'An toàn', icon: 'ms-Icon--Completed' };
  } catch (error) {
    console.error('Lỗi dự đoán:', error);
    return { status: 'Lỗi dự đoán', icon: 'ms-Icon--ErrorBadge' };
  }
}

// Hàm để cập nhật hiển thị kết quả ẩn (chỉ hiển thị URL, bỏ attachments)
function updateHiddenResults() {
  const resultDiv = document.getElementById("results-content");
  const loadingIndicator = document.getElementById("loading-indicator");
  const item = Office.context.mailbox.item;
  
  if (!item) {
    resultDiv.innerHTML = "<p class='ms-font-m'>Không có email nào được chọn.</p>";
    return;
  }

  // Hiển thị loading
  loadingIndicator.style.display = "block";
  resultDiv.innerHTML = "";

  // Lấy nội dung email dưới dạng HTML và trích xuất link (bỏ phần attachments)
  item.body.getAsync("html", (result) => {
    loadingIndicator.style.display = "none";
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const bodyHtml = result.value;
      const urls = extractUrlsFromHtml(bodyHtml);
      let urlList = "<h3>Các URL:</h3><ul class='ms-List'>";
      if (urls.length === 0) {
        urlList += "<li class='ms-ListItem'><i class='ms-Icon ms-Icon--Link'></i>Không tìm thấy Link trong email</li>";
      } else {
        urls.forEach(url => {
          urlList += `<li class='ms-ListItem'><i class='ms-Icon ms-Icon--Link'></i><a href="${url}" target="_blank">${url}</a></li>`;
        });
      }
      urlList += "</ul>";

      // Hiển thị kết quả (chỉ URL)
      resultDiv.innerHTML = urlList;
    } else {
      resultDiv.innerHTML = "<p class='ms-font-m ms-fontColor-error'>Lỗi khi lấy nội dung email: " + result.error.message + "</p>";
    }
  });
}

// Hàm để xử lý nhấp nút hiển thị tệp ẩn (với toggle)
function handleShowHidden() {
  const resultDisplay = document.getElementById("result-display");
  const buttonIcon = document.querySelector("#show-hidden-btn .ms-Button-icon i");
  const buttonLabel = document.querySelector("#show-hidden-btn .ms-Button-label");
  
  if (!isShowingHidden) {
    // Kích hoạt chế độ hiển thị
    isShowingHidden = true;
    resultDisplay.style.display = "block";
    buttonIcon.className = "ms-Icon ms-Icon--Hide";
    buttonLabel.textContent = "Ẩn";
    
    // Cập nhật kết quả ban đầu
    updateHiddenResults();
  } else {
    // Hủy kích hoạt chế độ
    isShowingHidden = false;
    resultDisplay.style.display = "none";
    buttonIcon.className = "ms-Icon ms-Icon--View";
    buttonLabel.textContent = "Xem";
  }
}

// Hàm để chạy kiểm tra lừa đảo AI (chỉ kiểm tra link, bỏ attachments)
async function runAiPhishingCheck() {
  const aiResultDiv = document.getElementById("ai-results-content");
  const aiLoadingIndicator = document.getElementById("ai-loading-indicator");
  const item = Office.context.mailbox.item;
  
  if (!item) {
    aiResultDiv.innerHTML = "<p class='ms-font-m'>Không có email nào được chọn.</p>";
    return;
  }

  // Hiển thị loading
  aiLoadingIndicator.style.display = "block";
  aiResultDiv.innerHTML = "";

  try {
    // Bỏ phần attachments hoàn toàn

    // Lấy URL từ nội dung email
    let urls = [];
    const bodyResult = await new Promise((resolve) => {
      item.body.getAsync("html", resolve);
    });
    if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
      urls = extractUrlsFromHtml(bodyResult.value);
    }

    // Phân tích AI chỉ cho URL
    let aiReport = "<h3>Phân tích Phishing bằng AI:</h3><ul class='ms-List'>";

    // Phân tích URL sử dụng predictUrl cục bộ
    for (const url of urls) {
      const pred = await predictUrl(url, bodyResult.value);
      
      // Thêm kiểu inline cho icon dựa trên trạng thái
      let iconStyle = '';
      if (pred.status === 'An toàn') {
        iconStyle = 'style="color: green;"';  // Xanh lá cây cho dấu tích (an toàn)
      } else if (pred.status === 'Nguy hiểm') {
        iconStyle = 'style="color: red;"';    // Đỏ cho dấu x/cảnh báo (nguy hiểm)
      } else {
        iconStyle = 'style="color: gray;"';   // Màu xám cho lỗi hoặc unknown (tùy chọn)
      }
      
      aiReport += `<li class='ms-ListItem'><i class='ms-Icon ${pred.icon}' ${iconStyle}></i>${url} - ${pred.status}.</li>`;
    }

    if (urls.length === 0) {
      aiReport += "<li class='ms-ListItem'><i class='ms-Icon ms-Icon--Info'></i>Không tìm thấy URL để phân tích.</li>";
    }

    aiReport += "</ul>";

    // Hiển thị kết quả AI
    aiResultDiv.innerHTML = aiReport;
  } catch (error) {
    aiResultDiv.innerHTML = "<p class='ms-font-m ms-fontColor-error'>Lỗi trong quá trình phân tích bằng AI: " + error.message + "</p>";
  } finally {
    aiLoadingIndicator.style.display = "none";
  }
}

// Hàm để xử lý nhấp nút kiểm tra AI (với toggle)
function handleAiCheck() {
  const aiResultDisplay = document.getElementById("ai-result-display");
  const buttonIcon = document.querySelector("#run-ai-check-btn .ms-Button-icon i");
  const buttonLabel = document.querySelector("#run-ai-check-btn .ms-Button-label");
  
  if (!isRunningAiCheck) {
    // Kích hoạt chế độ AI
    isRunningAiCheck = true;
    aiResultDisplay.style.display = "block";
    buttonIcon.className = "ms-Icon ms-Icon--Hide";
    buttonLabel.textContent = "Dừng";
    
    // Chạy kiểm tra AI ban đầu
    runAiPhishingCheck();
  } else {
    // Hủy kích hoạt chế độ
    isRunningAiCheck = false;
    aiResultDisplay.style.display = "none";
    buttonIcon.className = "ms-Icon ms-Icon--Robot";
    buttonLabel.textContent = "Chạy";
  }
}

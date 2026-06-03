<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
PageTitle = "聯絡我們 Contact Us"
Dim msgSent : msgSent = False
Dim msgError : msgError = ""

' Handle form submission
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
  Dim fName, fEmail, fPhone, fCompany, fPartno, fMsg, fSubject
  fName    = Trim(Request.Form("contact_name"))
  fEmail   = Trim(Request.Form("contact_email"))
  fPhone   = Trim(Request.Form("contact_phone"))
  fCompany = Trim(Request.Form("contact_company"))
  fPartno  = Trim(Request.Form("contact_partno"))
  fSubject = Trim(Request.Form("contact_subject"))
  fMsg     = Trim(Request.Form("contact_message"))

  If fName = "" Or fEmail = "" Or fMsg = "" Then
    msgError = "請填寫必填欄位（姓名、Email、訊息內容）。"
  Else
    ' Save to database
    On Error Resume Next
    Dim connMsg, sqlMsg
    Set connMsg = Server.CreateObject("ADODB.Connection")
    connMsg.Open "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("database/data_new.mdb")
    If Err.Number = 0 Then
      sqlMsg = "INSERT INTO messages (msg_name, msg_email, msg_phone, msg_company, msg_partno, msg_subject, msg_content, msg_date) VALUES ('" & Replace(fName,"'","''") & "','" & Replace(fEmail,"'","''") & "','" & Replace(fPhone,"'","''") & "','" & Replace(fCompany,"'","''") & "','" & Replace(fPartno,"'","''") & "','" & Replace(fSubject,"'","''") & "','" & Replace(fMsg,"'","''") & "',Now())"
      connMsg.Execute(sqlMsg)
      If Err.Number = 0 Then
        msgSent = True
      Else
        msgError = "儲存失敗，請直接寄信至 hch@hchtwn.com.tw"
      End If
      connMsg.Close: Set connMsg = Nothing
    Else
      msgError = "系統錯誤，請直接寄信至 hch@hchtwn.com.tw"
    End If
    On Error GoTo 0
  End If
End If
%>
<%@include file="inc_header.asp"%>

<!-- BREADCRUMB -->
<div class="breadcrumb">
  <div class="container">
    <ul class="breadcrumb-list">
      <li><a href="index.asp">首頁</a></li>
      <li>聯絡我們</li>
    </ul>
  </div>
</div>

<!-- PAGE HERO -->
<div class="page-hero">
  <div class="container">
    <h1><i class="fas fa-envelope" style="font-size:32px;margin-right:14px;opacity:0.8;"></i>聯絡我們 Contact Us</h1>
    <p>歡迎透過下列方式與我們聯繫，業務人員將盡快回覆您。</p>
  </div>
</div>

<section class="section">
  <div class="container">

    <!-- CONTACT OPTIONS CARDS -->
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:18px;margin-bottom:50px;">
      <div style="background:#fff;border:1px solid #E0E0E0;border-radius:10px;padding:24px;text-align:center;box-shadow:0 2px 10px rgba(0,0,0,0.06);transition:box-shadow 0.2s,transform 0.2s;" onmouseover="this.style.boxShadow='0 8px 30px rgba(0,0,0,0.12)';this.style.transform='translateY(-3px)'" onmouseout="this.style.boxShadow='0 2px 10px rgba(0,0,0,0.06)';this.style.transform=''">
        <div style="width:56px;height:56px;background:linear-gradient(135deg,#990000,#CC0000);border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 14px;font-size:24px;color:#fff;">
          <i class="fas fa-phone"></i>
        </div>
        <h4 style="font-weight:700;margin-bottom:6px;color:#111;">電話</h4>
        <a href="tel:+886223934278" style="font-size:15px;color:#CC0000;font-weight:600;">+886-2-2393-4278</a>
        <p style="font-size:12px;color:#888;margin-top:6px;">週一至週五 9:00–18:00</p>
      </div>
      <div style="background:#fff;border:1px solid #E0E0E0;border-radius:10px;padding:24px;text-align:center;box-shadow:0 2px 10px rgba(0,0,0,0.06);transition:box-shadow 0.2s,transform 0.2s;" onmouseover="this.style.boxShadow='0 8px 30px rgba(0,0,0,0.12)';this.style.transform='translateY(-3px)'" onmouseout="this.style.boxShadow='0 2px 10px rgba(0,0,0,0.06)';this.style.transform=''">
        <div style="width:56px;height:56px;background:linear-gradient(135deg,#002277,#003399);border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 14px;font-size:24px;color:#fff;">
          <i class="fas fa-envelope"></i>
        </div>
        <h4 style="font-weight:700;margin-bottom:6px;color:#111;">Email</h4>
        <a href="mailto:hch@hchtwn.com.tw" style="font-size:14px;color:#003399;font-weight:600;">hch@hchtwn.com.tw</a>
        <p style="font-size:12px;color:#888;margin-top:6px;">24 小時內回覆</p>
      </div>
      <div style="background:#E8FAF0;border:1.5px solid #06C755;border-radius:10px;padding:24px;text-align:center;box-shadow:0 2px 10px rgba(0,0,0,0.06);transition:box-shadow 0.2s,transform 0.2s;" onmouseover="this.style.boxShadow='0 8px 30px rgba(0,0,0,0.12)';this.style.transform='translateY(-3px)'" onmouseout="this.style.boxShadow='0 2px 10px rgba(0,0,0,0.06)';this.style.transform=''">
        <div style="width:56px;height:56px;background:#06C755;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 14px;font-size:24px;color:#fff;">
          <svg width="26" height="26" viewBox="0 0 24 24" fill="white"><path d="M19.952 12.375C19.952 8.33 15.9 5.03 11 5.03c-4.9 0-8.952 3.3-8.952 7.345 0 3.63 3.22 6.674 7.572 7.25.295.064.696.196.797.45.092.23.06.59.03.822 0 0-.107.639-.13.775-.04.23-.185.9.79.49.974-.408 5.254-3.094 7.17-5.297 1.32-1.452 1.673-2.904 1.673-4.49z"/></svg>
        </div>
        <h4 style="font-weight:700;margin-bottom:6px;color:#111;">LINE 官方帳號</h4>
        <a href="https://line.me/R/ti/p/@307puwgm" target="_blank" rel="noopener" style="font-size:14px;color:#06C755;font-weight:600;">@307puwgm</a>
        <p style="font-size:12px;color:#666;margin-top:6px;">即時回覆，方便快捷</p>
      </div>
      <div style="background:#fff;border:1px solid #E0E0E0;border-radius:10px;padding:24px;text-align:center;box-shadow:0 2px 10px rgba(0,0,0,0.06);transition:box-shadow 0.2s,transform 0.2s;" onmouseover="this.style.boxShadow='0 8px 30px rgba(0,0,0,0.12)';this.style.transform='translateY(-3px)'" onmouseout="this.style.boxShadow='0 2px 10px rgba(0,0,0,0.06)';this.style.transform=''">
        <div style="width:56px;height:56px;background:linear-gradient(135deg,#333,#555);border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 14px;font-size:24px;color:#fff;">
          <i class="fas fa-map-marker-alt"></i>
        </div>
        <h4 style="font-weight:700;margin-bottom:6px;color:#111;">地址</h4>
        <p style="font-size:13px;color:#555;line-height:1.7;">台灣台北市<br>Taipei, Taiwan R.O.C.</p>
      </div>
    </div>

    <!-- CONTACT GRID -->
    <div class="contact-grid">

      <!-- CONTACT FORM -->
      <div class="contact-form" style="background:#fff;border:1px solid #E0E0E0;border-radius:10px;padding:36px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">
        <h3><i class="fas fa-paper-plane" style="color:#CC0000;margin-right:10px;"></i>線上留言 Send Message</h3>
        <p style="font-size:13px;color:#888;margin-bottom:24px;">填寫以下表單，我們將盡快與您聯繫。標有 * 為必填。</p>

        <% If msgSent Then %>
        <div class="form-msg success" style="margin-bottom:20px;">
          <i class="fas fa-check-circle" style="margin-right:8px;"></i>
          <strong>留言已成功送出！</strong> 我們將於 24 小時內回覆您。謝謝！
        </div>
        <% ElseIf msgError <> "" Then %>
        <div class="form-msg error" style="margin-bottom:20px;">
          <i class="fas fa-exclamation-circle" style="margin-right:8px;"></i><%=msgError%>
        </div>
        <% End If %>

        <% If Not msgSent Then %>
        <form method="post" action="contact.asp" id="contactForm">
          <div class="form-row">
            <div class="form-group">
              <label>姓名 Name *</label>
              <input type="text" name="contact_name" placeholder="您的姓名" required>
            </div>
            <div class="form-group">
              <label>公司 Company</label>
              <input type="text" name="contact_company" placeholder="公司名稱">
            </div>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Email *</label>
              <input type="email" name="contact_email" placeholder="your@email.com" required>
            </div>
            <div class="form-group">
              <label>電話 Phone</label>
              <input type="tel" name="contact_phone" placeholder="+886-x-xxxx-xxxx">
            </div>
          </div>
          <div class="form-group">
            <label>詢問主題 Subject</label>
            <select name="contact_subject">
              <option value="產品詢問">產品詢問 Product Inquiry</option>
              <option value="報價">報價 Quotation</option>
              <option value="料號搜尋">料號搜尋 Part Number Search</option>
              <option value="技術支援">技術支援 Technical Support</option>
              <option value="其他">其他 Other</option>
            </select>
          </div>
          <div class="form-group">
            <label>詢問料號 Part No. (如有)</label>
            <input type="text" name="contact_partno" placeholder="e.g. STS-05, PTS-056...">
          </div>
          <div class="form-group">
            <label>訊息內容 Message *</label>
            <textarea name="contact_message" placeholder="請輸入您的訊息內容..." required></textarea>
          </div>
          <button type="submit" class="form-submit">
            <i class="fas fa-paper-plane" style="margin-right:8px;"></i>送出留言 Send Message
          </button>
        </form>
        <% End If %>
      </div>

      <!-- MAP + INFO -->
      <div>
        <div style="background:#fff;border:1px solid #E0E0E0;border-radius:10px;overflow:hidden;margin-bottom:24px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">
          <div style="padding:24px 28px;">
            <h3 style="font-size:20px;font-weight:700;margin-bottom:20px;color:#111;">
              <i class="fas fa-info-circle" style="color:#CC0000;margin-right:10px;"></i>聯絡資訊
            </h3>
            <div class="contact-row">
              <div class="ci"><i class="fas fa-building"></i></div>
              <div>
                <strong>公司名稱</strong>
                <p>鴻琪股份有限公司<br>HORNG CHIH HO CO., LTD.</p>
              </div>
            </div>
            <div class="contact-row">
              <div class="ci"><i class="fas fa-phone"></i></div>
              <div>
                <strong>電話 TEL</strong>
                <p><a href="tel:+886223934278">+886-2-2393-4278</a></p>
              </div>
            </div>
            <div class="contact-row">
              <div class="ci"><i class="fas fa-envelope"></i></div>
              <div>
                <strong>Email</strong>
                <p><a href="mailto:hch@hchtwn.com.tw">hch@hchtwn.com.tw</a></p>
              </div>
            </div>
            <div class="contact-row">
              <div class="ci"><i class="fas fa-globe"></i></div>
              <div>
                <strong>網站 Website</strong>
                <p><a href="https://www.hchtwn.com.tw" target="_blank">www.hchtwn.com.tw</a></p>
              </div>
            </div>
            <div class="contact-row">
              <div class="ci"><i class="fas fa-clock"></i></div>
              <div>
                <strong>營業時間 Business Hours</strong>
                <p>週一至週五 Mon–Fri: 9:00–18:00<br>週六、日及國定假日休息</p>
              </div>
            </div>
          </div>
        </div>

        <!-- LINE WIDGET -->
        <div style="background:linear-gradient(135deg,#05a848,#06C755);border-radius:10px;padding:24px;color:#fff;text-align:center;box-shadow:0 4px 20px rgba(6,199,85,0.3);">
          <svg width="44" height="44" viewBox="0 0 24 24" fill="white" style="margin-bottom:12px;"><path d="M19.952 12.375C19.952 8.33 15.9 5.03 11 5.03c-4.9 0-8.952 3.3-8.952 7.345 0 3.63 3.22 6.674 7.572 7.25.295.064.696.196.797.45.092.23.06.59.03.822 0 0-.107.639-.13.775-.04.23-.185.9.79.49.974-.408 5.254-3.094 7.17-5.297 1.32-1.452 1.673-2.904 1.673-4.49z"/></svg>
          <h4 style="font-size:18px;font-weight:700;margin-bottom:8px;">LINE 官方帳號</h4>
          <p style="font-size:13px;opacity:0.9;margin-bottom:16px;">加入好友，即時與業務人員溝通洽詢</p>
          <div style="background:rgba(255,255,255,0.2);display:inline-block;padding:4px 16px;border-radius:4px;font-size:13px;font-weight:600;margin-bottom:16px;">@307puwgm</div>
          <br>
          <a href="https://line.me/R/ti/p/@307puwgm" target="_blank" rel="noopener" style="display:inline-block;background:#fff;color:#06C755;padding:12px 32px;border-radius:6px;font-size:15px;font-weight:700;">加入 LINE 好友</a>
        </div>
      </div>
    </div>
  </div>
</section>

<%@include file="inc_footer.asp"%>

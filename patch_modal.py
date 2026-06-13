"""Patch HTML template to add confirmation modal for missing bureau data."""
from pathlib import Path

html_path = Path(__file__).parent / "web" / "templates" / "index.html"
html = html_path.read_text(encoding="utf-8")

# 1. Add modal CSS before </style>
modal_css = """
.modal-overlay{position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);
               display:none;justify-content:center;align-items:center;z-index:1000}
.modal{background:#22242e;border:1px solid #2e3040;border-radius:12px;padding:28px 32px;
       max-width:420px;width:90%;text-align:center}
.modal h3{font-size:16px;color:#ffc845;margin-bottom:12px}
.modal p{font-size:14px;color:#9a9eb8;margin-bottom:16px;line-height:1.5}
.modal .btn-row{display:flex;gap:12px;justify-content:center}
"""
if "modal-overlay" not in html:
    html = html.replace("</style>", modal_css + "\n</style>")

# 2. Add modal HTML after log_box
modal_html = """
  <div class="modal-overlay" id="modal_overlay">
    <div class="modal">
      <h3>\u7f3a\u5931\u767b\u8bb0\u673a\u5173\u6570\u636e</h3>
      <p id="modal_msg"></p>
      <div class="btn-row">
        <button class="btn btn-primary" onclick="confirmSupplement()">\u786e\u8ba4\u8865\u5145</button>
        <button class="btn btn-secondary" onclick="cancelSupplement()">\u53d6\u6d88</button>
      </div>
    </div>
  </div>"""
if "modal_overlay" not in html:
    html = html.replace(
        '<div class="log-box" id="log_box"></div>',
        '<div class="log-box" id="log_box"></div>\n' + modal_html,
    )

# 3. Add modal JS functions before doPrecheck
modal_js = """
function showModal(msg){document.getElementById('modal_msg').textContent=msg;document.getElementById('modal_overlay').style.display='flex';}
function hideModal(){document.getElementById('modal_overlay').style.display='none';}
function confirmSupplement(){
  hideModal();
  var saveDir=document.getElementById('output_dir').value||'';
  fetch('/api/export_template',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({save_dir:saveDir})})
  .then(function(r){return r.json()}).then(function(d){
    if(d.ok){fetch('/api/confirm_supplement',{method:'POST'}).then(function(r2){return r2.json()}).then(function(){pollState()});}
    else if(d.error){alert(d.error);}
  });
}
function cancelSupplement(){hideModal();fetch('/api/cancel_supplement',{method:'POST'}).then(function(r){return r.json()}).then(function(){pollState()});}

"""
if "confirmSupplement" not in html:
    html = html.replace("function doPrecheck()", modal_js + "function doPrecheck()")

# 4. Update updateUI to handle needs_confirm
old_status_line = "  $('status').textContent=d.status;\n  $('status').className='status status-'+d.status_color;"
new_status_block = (
    "  $('status').textContent=d.status==='needs_confirm'?'\u7f3a\u5931\u6570\u636e\u8bf7\u786e\u8ba4':d.status;\n"
    "  $('status').className='status status-'+d.status_color;\n"
    "  if(d.status==='needs_confirm'&&d.missing_count>0){showModal('\u7f3a\u5931 '+d.missing_count+' \u6761\u4f01\u4e1a\u767b\u8bb0\u673a\u5173\u6570\u636e\\n\u70b9\u51fb\u300c\u786e\u8ba4\u8865\u5145\u300d\u751f\u6210\u6a21\u677f\\n\u70b9\u51fb\u300c\u53d6\u6d88\u300d\u7ee7\u7eed\u7559\u5728\u6b64\u6b65\u9aa4');}"
)
if "needs_confirm" not in html:
    html = html.replace(old_status_line, new_status_block)

html_path.write_text(html, encoding="utf-8")
print("HTML patched with modal dialog")

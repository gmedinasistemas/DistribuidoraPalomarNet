using Microsoft.Web.WebView2.WinForms;
using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet
{
    /// <summary>
    /// Formulario de progreso para el recálculo de precios al salir de la configuración por familia.
    /// Utiliza un control <see cref="WebView2"/> para mostrar una interfaz HTML moderna
    /// con cards e iconos para cada paso del proceso.
    /// </summary>
    public partial class frmProgresoRecalculo : Form
    {
        private WebView2 _wv;
        private bool _ready = false;

        public frmProgresoRecalculo()
        {
            InitializeComponent();

            _wv = new WebView2();
            _wv.Dock = DockStyle.Fill;
            _wv.DefaultBackgroundColor = System.Drawing.Color.FromArgb(15, 17, 23);
            Controls.Add(_wv);

            Load += async (s, e) => await InitWebViewAsync();
        }

        private async Task InitWebViewAsync()
        {
            try
            {
                await _wv.EnsureCoreWebView2Async(null);
                _wv.CoreWebView2.Settings.IsStatusBarEnabled = false;
                _wv.CoreWebView2.Settings.AreDevToolsEnabled = false;
                _wv.CoreWebView2.Settings.IsZoomControlEnabled = false;
                _wv.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;
                _wv.NavigateToString(GetHtml());
                _wv.NavigationCompleted += (s, e) => { _ready = true; };
            }
            catch
            {
                _ready = false;
            }
        }

        /// <summary>
        /// Espera hasta que el WebView2 termine de cargar el HTML inicial.
        /// Bombea la cola de mensajes de Windows Forms para no bloquear el hilo UI.
        /// </summary>
        public void WaitReady(int timeoutMs = 5000)
        {
            var sw = Stopwatch.StartNew();
            while (!_ready && sw.ElapsedMilliseconds < timeoutMs)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(30);
            }
        }

        /// <summary>Cambia el estado visual de un paso del proceso.</summary>
        /// <param name="index">Índice del paso (0–3).</param>
        /// <param name="estado">"active" | "done" | "error" | "pending"</param>
        public void SetStep(int index, string estado)
        {
            if (!_ready) return;
            _ = _wv.ExecuteScriptAsync($"setStep({index},'{estado}')");
            Application.DoEvents();
        }

        /// <summary>Actualiza el texto de la barra de mensaje inferior.</summary>
        public void SetMensaje(string texto)
        {
            if (!_ready) return;
            var esc = (texto ?? "").Replace("\\", "\\\\").Replace("'", "\\'");
            _ = _wv.ExecuteScriptAsync($"setMensaje('{esc}')");
            Application.DoEvents();
        }

        /// <summary>Activa el estado de éxito en la barra de mensaje (punto verde).</summary>
        public void SetFinalizado()
        {
            if (!_ready) return;
            _ = _wv.ExecuteScriptAsync("setFinalizado()");
            Application.DoEvents();
        }

        /// <summary>Activa el estado de error en la barra de mensaje (punto rojo).</summary>
        public void SetError(string msg)
        {
            if (!_ready) return;
            var esc = (msg ?? "Error en el proceso").Replace("\\", "\\\\").Replace("'", "\\'");
            _ = _wv.ExecuteScriptAsync($"setError('{esc}')");
            Application.DoEvents();
        }

        private static string GetHtml() => """
<!DOCTYPE html>
<html lang='es'>
<head>
<meta charset='UTF-8'>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{
  font-family:'Segoe UI',system-ui,sans-serif;
  background:linear-gradient(160deg,#0f1117 0%,#1a1a2e 55%,#16213e 100%);
  min-height:100vh;
  display:flex;flex-direction:column;align-items:center;
  padding:22px 20px 16px;
  color:#e2e8f0;overflow:hidden;
}
.header{text-align:center;margin-bottom:18px}
.logo{
  width:48px;height:48px;
  background:linear-gradient(135deg,#0ea5e9,#6366f1);
  border-radius:14px;display:flex;align-items:center;justify-content:center;
  margin:0 auto 10px;
  box-shadow:0 8px 24px rgba(99,102,241,.4);
}
h1{font-size:15px;font-weight:700;color:#f1f5f9;letter-spacing:-.2px;margin-bottom:3px}
.sub{font-size:11px;color:#475569}
.steps{width:100%;max-width:420px;display:flex;flex-direction:column;gap:8px}
.step{
  background:rgba(255,255,255,.04);
  border:1px solid rgba(255,255,255,.07);
  border-radius:10px;padding:11px 14px;
  display:flex;align-items:center;gap:12px;
  transition:background .3s,border-color .3s,box-shadow .3s;
}
.step.active{
  background:rgba(99,102,241,.13);
  border-color:rgba(99,102,241,.5);
  box-shadow:0 0 20px rgba(99,102,241,.2);
}
.step.done{background:rgba(16,185,129,.08);border-color:rgba(16,185,129,.32)}
.step.error{background:rgba(239,68,68,.09);border-color:rgba(239,68,68,.38)}
.ic{
  width:34px;height:34px;border-radius:50%;
  display:flex;align-items:center;justify-content:center;flex-shrink:0;
}
.step.pending .ic{background:rgba(255,255,255,.05);border:1.5px solid rgba(255,255,255,.09);color:#475569}
.step.active  .ic{background:rgba(99,102,241,.28);border:1.5px solid #6366f1;color:#a5b4fc}
.step.done    .ic{background:rgba(16,185,129,.2);border:1.5px solid #10b981;color:#10b981}
.step.error   .ic{background:rgba(239,68,68,.2);border:1.5px solid #ef4444;color:#ef4444}
.spin{
  width:16px;height:16px;
  border:2px solid rgba(99,102,241,.25);border-top-color:#6366f1;
  border-radius:50%;animation:sp .75s linear infinite;
}
@keyframes sp{to{transform:rotate(360deg)}}
.inf{flex:1}
.lbl{font-size:10px;color:#475569;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px}
.ttl{font-size:13px;font-weight:600}
.step.pending .ttl{color:#475569}
.step.active  .ttl{color:#c7d2fe}
.step.done    .ttl{color:#6ee7b7}
.step.error   .ttl{color:#fca5a5}
.msgbar{
  margin-top:14px;
  background:rgba(255,255,255,.04);
  border:1px solid rgba(255,255,255,.07);
  border-radius:8px;padding:10px 14px;
  width:100%;max-width:420px;
  display:flex;align-items:center;gap:10px;
}
.dot{
  width:7px;height:7px;border-radius:50%;background:#6366f1;flex-shrink:0;
  animation:pu 1.4s ease-in-out infinite;
}
@keyframes pu{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.4;transform:scale(.65)}}
#msg{font-size:12px;color:#94a3b8}
</style>
</head>
<body>

<div class='header'>
  <div class='logo'>
    <svg width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'>
      <circle cx='12' cy='12' r='3'></circle>
      <path d='M19.07 4.93a10 10 0 0 1 0 14.14M4.93 4.93a10 10 0 0 0 0 14.14'></path>
    </svg>
  </div>
  <h1>Aplicando Configuraci&#243;n por Familia</h1>
  <p class='sub'>Palomar Distribuidora &mdash; Recalculando precios</p>
</div>

<div class='steps' id='sc'></div>

<div class='msgbar'>
  <div class='dot' id='dot'></div>
  <span id='msg'>Iniciando proceso...</span>
</div>

<script>
var S=[
  {t:'Aplicando descuentos del cliente',
   i:'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>'},
  {t:'Recalculando precios generales',
   i:'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"></line><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"></path></svg>'},
  {t:'Recalculando precios de ofertas',
   i:'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z"></path><line x1="7" y1="7" x2="7.01" y2="7"></line></svg>'},
  {t:'Actualizando cat&#225;logo',
   i:'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"></polyline><polyline points="23 20 23 14 17 14"></polyline><path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4l-4.64 4.36A9 9 0 0 1 3.51 15"></path></svg>'}
];

var sc=document.getElementById('sc');
S.forEach(function(s,i){
  var d=document.createElement('div');
  d.className='step pending';
  d.id='s'+i;
  d.innerHTML='<div class="ic" id="ic'+i+'">'+s.i+'</div>'
    +'<div class="inf"><div class="lbl">Paso '+(i+1)+' de '+S.length+'</div>'
    +'<div class="ttl">'+s.t+'</div></div>';
  sc.appendChild(d);
});

var ICO_CHECK='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>';
var ICO_X='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>';

function setStep(i,e){
  var el=document.getElementById('s'+i);
  var ic=document.getElementById('ic'+i);
  if(!el) return;
  el.className='step '+e;
  if(e==='active') ic.innerHTML='<div class="spin"></div>';
  else if(e==='done')  ic.innerHTML=ICO_CHECK;
  else if(e==='error') ic.innerHTML=ICO_X;
  else ic.innerHTML=S[i].i;
}

function setMensaje(t){
  document.getElementById('msg').textContent=t;
}

function setFinalizado(){
  var d=document.getElementById('dot');
  d.style.background='#10b981';
  d.style.animation='none';
  document.getElementById('msg').style.color='#6ee7b7';
}

function setError(t){
  var d=document.getElementById('dot');
  d.style.background='#ef4444';
  d.style.animation='none';
  var m=document.getElementById('msg');
  m.style.color='#fca5a5';
  if(t) m.textContent=t;
}
</script>
</body>
</html>
""";
    }
}

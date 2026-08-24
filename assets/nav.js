(function(){
  const NAV_ITEMS = [
    { href: "index.html", label: "Score", icon: "\u25C6" },
    { href: "man-status.html", label: "Man Status", icon: "\u2694" },
    { href: "draft-classes.html", label: "Classes", icon: "\u25A6" },
    { href: "roster.html", label: "Roster", icon: "\u2261" },
    { href: "legacy.html", label: "Legacy", icon: "\u265A" },
    { href: "records.html", label: "Records", icon: "\u2691" },
    { href: "analytics.html", label: "Stats", icon: "\u2726" },
    { href: "wildcards.html", label: "Wild", icon: "\u2668" },
    { href: "rules.html", label: "Rules", icon: "\u25CE" },
  ];

  const style = document.createElement('style');
  style.textContent = `
    .app-nav{
      position:fixed; left:0; right:0; bottom:0; z-index:50;
      display:flex;
      background:rgba(11,13,18,0.94);
      backdrop-filter:blur(10px);
      border-top:1px solid #252C38;
      padding:6px 4px calc(6px + env(safe-area-inset-bottom));
    }
    .app-nav a{
      flex:1;
      display:flex; flex-direction:column; align-items:center; gap:1px;
      text-decoration:none;
      color:#8B93A1;
      font-size:8.4px;
      letter-spacing:0.02em;
      text-transform:uppercase;
      padding:5px 1px;
      border-radius:9px;
      min-width:0;
    }
    .app-nav a .ic{font-size:14px; line-height:1;}
    .app-nav a.active{color:#EDEFF2;}
    .app-nav a.active .ic{color:#F2A93B;}
  `;
  document.head.appendChild(style);

  const current = (location.pathname.split('/').pop() || 'index.html');

  const nav = document.createElement('nav');
  nav.className = 'app-nav';
  nav.innerHTML = NAV_ITEMS.map(item => `
    <a href="${item.href}" class="${item.href===current ? 'active' : ''}">
      <span class="ic">${item.icon}</span>
      <span>${item.label}</span>
    </a>
  `).join('');

  document.body.appendChild(nav);
})();

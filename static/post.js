/* Renders a /writing/<slug>/ page: Markdown from <script id="md">, title + date + prev/next from /writing/posts.json. */
(async () => {
  const md = document.getElementById("md");
  if (md && window.marked) document.getElementById("post").innerHTML = marked.parse(md.textContent.replace(/^\s*\n/, ""));

  let posts = [];
  try { posts = await (await fetch("/writing/posts.json", { cache: "no-cache" })).json(); } catch (e) { return; }
  posts.sort((a, b) => b.date.localeCompare(a.date));
  const slug = location.pathname.replace(/\/(index\.html)?$/, "").split("/").pop();
  const i = posts.findIndex(p => p.slug === slug);
  if (i < 0) return;

  const p = posts[i];
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const fmt = d => new Date(d + "T00:00:00Z").toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric", timeZone: "UTC" });
  document.title = p.title + " — Bashid Effendi";
  document.getElementById("title").textContent = p.title;
  const t = document.getElementById("date");
  t.dateTime = p.date;
  t.textContent = fmt(p.date);

  const newer = posts[i - 1], older = posts[i + 1];
  document.getElementById("pn").innerHTML =
    (newer ? `<a href="/writing/${esc(newer.slug)}/"><span aria-hidden="true">←</span><span class="ttl">${esc(newer.title)}</span></a>` : "") +
    (older ? `<a class="next" href="/writing/${esc(older.slug)}/"><span class="ttl">${esc(older.title)}</span><span aria-hidden="true">→</span></a>` : "");
})();

# Docs articles

One file per article. The file name is the article's address: `quick-start.html` is shown at `/docs/quick-start`.

Each file starts with a short header, then the article itself:

```
<!--
section: get started        (the menu section it sits in)
label: quick start          (its name in the menu)
title: Quick start          (the big heading)
intro: One or two sentences under the heading.
order: 2                    (position in the menu; previous/next follow it)
icon: bolt                  (menu icon: lines, bolt, eye, download, compass, shield, lock, page, keyboard)
badge: soon                 (optional small tag in the menu)
-->
```

To change the wording, edit the text and leave the tags (`<p>`, `<h2>`, `<li>` and so on) in place.
The menu, previous/next and "on this page" are built from these files: "on this page" lists every `<h2>`
(or, if a heading has `data-toc="Shorter name"`, that shorter name).

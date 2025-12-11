# ACL SRW 2026 Website

Static site for the ACL SRW 2026 workshop, built with [Jekyll](https://jekyllrb.com/) and deployed via GitHub Pages.

## Project structure

- `_layouts/`, `_includes/`, `assets/`, `css/`, `images/`: shared templates, partials, and static assets.
- `_pages/`: all standalone Markdown pages (author guidelines, sponsors, etc.). Each file includes YAML front matter describing its layout, permalink, and metadata.
- `index.md`: landing page rendered at `/`.
- `_config.yml`: global site configuration, including the `pages` collection, navigation behavior, and metadata.

## Working with pages

1. Create or update Markdown files inside `_pages/`.
2. Include front matter such as:

   ```yaml
   ---
   layout: page
   title: "Author Guidelines"
   permalink: /author
   nav_hidden: false
   published: true
   order: 2
   ---
   ```

3. `permalink` controls the final URL. `order` determines the position within the navigation bar.
4. Set `nav_hidden: true` to keep a page published but off the top navigation. Set `published: false` to remove it from the build entirely.
5. Push your changes to the repository. GitHub Pages automatically rebuilds the site a minute or two after each push.

## Local development

```bash
bundle install          # first time only
bundle exec jekyll serve
```

Visit `http://localhost:4000` to preview the site. Press `Ctrl+C` to stop the development server.

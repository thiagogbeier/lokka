// @ts-check
// `@type` JSDoc annotations allow editor autocompletion and type checking
// (when paired with `@ts-check`).
// There are various equivalent ways to declare your Docusaurus config.
// See: https://docusaurus.io/docs/api/docusaurus-config

import {themes as prismThemes} from 'prism-react-renderer';

// This runs in Node.js - Don't use client-side code here (browser APIs, JSX...)

/** @type {import('@docusaurus/types').Config} */
const config = {
  title: 'Lokka',
  tagline: 'Beyond Commands, Beyond Clicks. A glimpse into the future of managing Microsoft 365 with AI!',
  favicon: 'img/favicon.ico',

  // Set the production url of your site here
  url: 'https://lokka.dev',
  // Set the /<baseUrl>/ pathname under which your site is served
  // For GitHub pages deployment, it is often '/<projectName>/'
  baseUrl: '/',

  // GitHub pages deployment config.
  // If you aren't using GitHub pages, you don't need these.
  organizationName: 'merill', // Usually your GitHub org/user name.
  projectName: 'lokka', // Usually your repo name.

  onBrokenLinks: 'throw',
  onBrokenMarkdownLinks: 'warn',

  // Even if you don't use internationalization, you can use this field to set
  // useful metadata like html lang. For example, if your site is Chinese, you
  // may want to replace "en" with "zh-Hans".
  i18n: {
    defaultLocale: 'en',
    locales: ['en'],
  },

  presets: [
    [
      'classic',
      /** @type {import('@docusaurus/preset-classic').Options} */
      ({
        docs: {
          sidebarPath: './sidebars.js',
          // Please change this to your repo.
          // Remove this to remove the "edit this page" links.
          editUrl:
            'https://github.com/merill/lokka/tree/main/website/docs',
        },
        blog: {
          showReadingTime: true,
          feedOptions: {
            type: ['rss', 'atom'],
            xslt: true,
          },
          // Please change this to your repo.
          // Remove this to remove the "edit this page" links.
          editUrl:
            'https://github.com/merill/lokka/tree/main/docs/',
          // Useful options to enforce blogging best practices
          onInlineTags: 'warn',
          onInlineAuthors: 'warn',
          onUntruncatedBlogPosts: 'warn',
        },
        theme: {
          customCss: './src/css/custom.css',
        },
      }),
    ],
  ],

  themeConfig:
    /** @type {import('@docusaurus/preset-classic').ThemeConfig} */
    ({
      // Replace with your project's social card
      image: 'img/docusaurus-social-card.png',
      navbar: {
        title: 'Lokka',
        logo: {
          alt: 'Lokka logo',
          src: 'img/logo.svg',
        },
        items: [
          {
            type: "docSidebar",
            sidebarId: "siteSidebar",
            position: "left",
            label: "Docs",
          },
          {
            href: 'https://graphxray.merill.net',
            label: 'Graph X-Ray',
            position: 'right',
          },
          {
            href: 'https://cmd.ms',
            label: 'cmd.ms',
            position: 'right',
          },
          {
            href: 'https://entra.news',
            label: 'entra.news',
            position: 'right',
          },
          {
            href: 'https://entra.chat',
            label: 'entra.chat',
            position: 'right',
          },
          {
            href: 'https://merill.net',
            label: 'merill.net',
            position: 'right',
          },
          {
            "aria-label": "GitHub Repository",
            className: "navbar--github-link",
            href: "https://github.com/merill/lokka",
            position: "right",
          },
        ],
      },
      footer: {
        style: 'dark',
        links: [
          {
            title: 'Learn',
            items: [
              {
                label: 'About me',
                href: 'https://merill.net/about',
              },
            ],
          },
          {
            title: 'Community',
            items: [
              {
                label: 'X',
                href: 'https://x.com/merill',
              },
              {
                label: 'Bluesky',
                href: 'https://bsky.app/profile/merill.net',
              },
            ],
          },
          {
            title: 'More',
            items: [
              {
                label: 'Blog',
                href: 'https://merill.net',
              },
              {
                label: 'GitHub',
                href: 'https://github.com/merill/lokka',
              },
            ],
          },
        ],
        copyright: `Copyright © ${new Date().getFullYear()} Merill Fernando.`,
      },
      prism: {
        theme: prismThemes.github,
        darkTheme: prismThemes.dracula,
      },
    }),
};

export default config;

import useDocusaurusContext from "@docusaurus/useDocusaurusContext";
import Layout from "@theme/Layout";

export default function Home() {
  const { siteConfig } = useDocusaurusContext();
  return (
    <Layout
      title={`${siteConfig.title}`}
      description="Beyond Commands, Beyond Clicks. A glimpse into the future of Managing Microsoft 365!">
      <main>
        <div className="hero hero--primary">
          <div className="container">
            <h1 className="hero__title">{siteConfig.title}</h1>
            <p className="hero__subtitle">{siteConfig.tagline}</p>
          </div>
        </div>
      </main>
    </Layout>
  );
}

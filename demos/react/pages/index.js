import Head from 'next/head';

export default function Index() { return (
<div>
  <Head>
    <meta httpEquiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>SheetJS Next.JS Demo</title>
    <script src="/shim.js"></script>
    <style jsx>{`
      body, #app { height: 100%; };
    `}</style>
  </Head>
  <pre>
<h3>SheetJS Next.JS Demos</h3>
All demos read from /public/sheetjs.xlsx.<br/>
<br/>
- <a href="/getStaticProps">getStaticProps</a><br/>
<br/>
- <a href="/getServerSideProps">getServerSideProps</a><br/>
<br/>
- <a href="/getStaticPaths">getStaticPaths</a><br/>
  </pre>
</div>
); }
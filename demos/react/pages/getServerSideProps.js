import Head from 'next/head';
import { readFile, utils } from 'xlsx';
import { join } from 'path';
import { cwd } from 'process';

export default function Index({html, type}) { return (
<div>
  <Head>
    <meta httpEquiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>SheetJS Next.JS {type} Demo</title>
    <script src="/shim.js"></script>
    <style jsx>{`
      body, #app { height: 100%; };
    `}</style>
  </Head>
  <pre>
<h3>SheetJS Next.JS {type} Demo</h3>
This demo reads from /public/sheetjs.xlsx and generates HTML from the first sheet.
  </pre>
  <div dangerouslySetInnerHTML={{ __html: html }} />
</div>
); }

export async function getServerSideProps() {
  const wb = readFile(join(cwd(), "public", "sheetjs.xlsx"))
  return {
    props: {
      type: "getStaticProps",
      html: utils.sheet_to_html(wb.Sheets[wb.SheetNames[0]]),
    },
  }
}
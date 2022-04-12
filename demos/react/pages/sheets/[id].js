import Head from 'next/head';
import { readFile, utils } from 'xlsx';
import { join } from 'path';
import { cwd } from 'process';

export default function Index({html, type, name}) { return (
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
This demo reads from /public/sheetjs.xlsx.<br/>
<br/>
<b>{name}</b>
  </pre>
  <div dangerouslySetInnerHTML={{ __html: html }} />
</div>
); }

let cache = [];

export async function getStaticProps(ctx) {
  if(!cache || !cache.length) {
    const wb = readFile(join(cwd(), "public", "sheetjs.xlsx"));
    cache = wb.SheetNames.map((name) => ({ name, sheet: wb.Sheets[name] }));
  }
  const entry = cache[ctx.params.id];
  return {
    props: {
      type: "getStaticPaths",
      name: entry.name,
      id: ctx.params.id.toString(),
      html: entry.sheet ? utils.sheet_to_html(entry.sheet) : "",
    },
  }
}

export async function getStaticPaths() {
  const wb = readFile(join(cwd(), "public", "sheetjs.xlsx"));
  cache = wb.SheetNames.map((name) => ({ name, sheet: wb.Sheets[name] }));
  return {
    paths: wb.SheetNames.map((name, idx) => ({ params: { id: idx.toString()  } })),
    fallback: false,
  };
}

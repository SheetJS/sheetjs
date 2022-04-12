import Head from 'next/head';
import Link from "next/link";
import { readFile, utils } from 'xlsx';
import { join } from 'path';
import { cwd } from 'process';

export default function Index({snames, type}) { return (
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
This demo reads from /public/sheetjs.xlsx.  Each worksheet maps to a path:<br/><br/>
{snames.map((sname, idx) => (<>
  <Link key={idx} href="/sheets/[id]" as={`/sheets/${idx}`}><a>{`Sheet index=${idx} name="${sname}"`}</a></Link>
  <br/>
  <br/>
</>))}

  </pre>
</div>
); }

export async function getStaticProps() {
  const wb = readFile(join(cwd(), "public", "sheetjs.xlsx"))
  return {
    props: {
      type: "getStaticPaths",
      snames: wb.SheetNames,
    },
  }
}
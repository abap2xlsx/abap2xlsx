import {initializeABAP} from "../output/init.mjs";
import {cl_excel_test} from "../output/cl_excel_test.clas.mjs";
import * as fs from "node:fs";

await initializeABAP();

const test = new cl_excel_test();
const buf = Buffer.from((await test.run()).get().toLowerCase(), "hex");
fs.writeFileSync("foo.xlsx", buf);
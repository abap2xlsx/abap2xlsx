import {initializeABAP} from "../output/init.mjs";
import {cl_excel_test} from "../output/cl_excel_test.clas.mjs";

await initializeABAP();

const test = new cl_excel_test();
console.dir(await test.run());
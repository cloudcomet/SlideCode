const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
const fn = new AsyncFunction('console.log("start"); await new Promise(r => setTimeout(r, 100)); console.log("end");');
fn().then(() => console.log("done"));

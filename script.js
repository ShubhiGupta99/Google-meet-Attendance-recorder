
const puppy =require("puppeteer");
const xlsx=require("xlsx");

const meetLink=process.argv[2]; //"https://meet.google.com/pbn-evak-dia"
const meetDuration=process.argv[3]*60*1000;

async function main()
{
    const browser=await puppy.launch({
        headless:false,
        defaultViewPort:null,
        args: ['--disable-web-security', '--user-data-dir', '--allow-running-insecure-content' ]
      
      }); 
      
      var strtTime = Date();
      var date=new Date().toDateString();
      console.log(strtTime);
      console.log(date);
       

      let pages=await browser.pages();
      let page=pages[0];
      await page.setViewport({width: 1533, height: 705});
      await page.goto(meetLink);
      
      await page.waitForSelector(".U26fgb.JRY2Pb.mUbCce.kpROve.uJNmj.QmxbVb.HNeRed[aria-label='Turn off microphone (CTRL + D)'] .IYwVEf.HotEze.uB7U9e.nAZzG");
      await page.click(".U26fgb.JRY2Pb.mUbCce.kpROve.uJNmj.QmxbVb.HNeRed[aria-label='Turn off microphone (CTRL + D)'] .IYwVEf.HotEze.uB7U9e.nAZzG");
      await page.click(".U26fgb.JRY2Pb.mUbCce.kpROve.uJNmj.QmxbVb.HNeRed[aria-label='Turn off camera (CTRL + E)'] .IYwVEf.HotEze.nAZzG");
      
      await page.waitForSelector(".uArJ5e.UQuaGc.Y5sE8d.uyXBBb.xKiqt",{visible:true});
      await page.click(".uArJ5e.UQuaGc.Y5sE8d.uyXBBb.xKiqt .NPEfkd.RveJvd.snByac"); 

      await wait(300000);//wait 5 mins
      
      let names1=await participants(page); 
      
      await wait(meetDuration/2-300000);
      let namesAtHalf=await participants(page);

      await wait(meetDuration-300000-meetDuration/2);
      
      let names2=await participants(page);
      
      let finalList=[];
      let presentAtFirstHalf=names1.filter(x => namesAtHalf.includes(x));
      let presentAtSecHalf=names2.filter(x => namesAtHalf.includes(x));

      let others1=names1.filter(x=>!namesAtHalf.includes(x));
      let others2=names2.filter(x=>!namesAtHalf.includes(x));
      let others_3=namesAtHalf.filter(x=>!names1.includes(x));
      let others3=others_3.filter(x=>!names2.includes(x));

      let leftEarly=presentAtFirstHalf.filter(x => !names2.includes(x));

      let presentAll=presentAtFirstHalf.filter(x => names2.includes(x)); //participants present all time

      let lateParticipants=presentAtSecHalf.filter(x => !names1.includes(x));
      
      for(let i of others1)
      {
        finalList.push([i,"-"]);
      }
      for(let i of others2)
      {
        finalList.push([i,"-"]);
      }
      for(let i of others3)
      {
        finalList.push([i,"-"]);
      }
      
      for(let i of leftEarly)
      {
        finalList.push([ i, "Left Early" ]);
      }

      for(let i of lateParticipants)
      {
        finalList.push([ i, "Joined Late" ]);
      }

      for(let i of presentAll)
      {
        finalList.push([ i, "Present" ]);
      }
    
      finalList.unshift(["PARTICIPANTS ","ATTENDANCE "]);
       
      const wb=xlsx.utils.book_new();
      const ws=xlsx.utils.aoa_to_sheet(finalList);
      ws['!cols'] = fitToColumn(finalList);

      function fitToColumn(finalList) {
      return finalList[0].map((a, i) => ({ wch: Math.max(...finalList.map(a2 => a2[i].toString().length)) }));
      }

      xlsx.utils.book_append_sheet(wb,ws);
      xlsx.writeFile(wb,"Attendance Report(" +date+ ").xlsx",{
        headerFooter:{firstHeader: " Meeting link : "+meetLink+"on "+date}});  
      
      console.log("Attendance recorded");  

      await browser.close();

 }

async function participants(page)
{ 
  try{
       await page.waitForSelector(".uArJ5e.UQuaGc.kCyAyd.QU4Gid.foXzLb.IeuGXd .Hdh4hc.cIGbvc.NMm5M",{visible:true});
       await page.click(".uArJ5e.UQuaGc.kCyAyd.QU4Gid.foXzLb.IeuGXd .Hdh4hc.cIGbvc.NMm5M");
       return participantsBox(page);
    }
  catch(error)
    {
      return participantsBox(page);
    }
  
} 

async function participantsBox(page)
{  let names=[];
   let nameSpan=await page.$$("div[role='listitem'] .ZjFb7c",{visible:true});

   for(let i=1;i<nameSpan.length;i++)
    {
     let name= await page.evaluate(function(ele)
     {
     return ele.textContent; 
     },nameSpan[i]);
     names.push(name);
    } 
  return names;
}

async function wait(t)
{
    return new Promise(function(resolve,reject){
        setTimeout(function(){
            resolve();},t);
        })
}


main();
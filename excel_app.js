// excel_app.js
const video = document.getElementById('video');
const canvas = document.getElementById('capture');
const labelInput = document.getElementById('label');
const priceInput = document.getElementById('price');
const addBtn = document.getElementById('add-example');
const predictBtn = document.getElementById('predict-btn');
const statusEl = document.getElementById('status');
const resultCard = document.getElementById('result');
const predLabel = document.getElementById('pred-label');
const predPrice = document.getElementById('pred-price');
const predProb = document.getElementById('pred-prob');
const exportBtn = document.getElementById('export-excel');

let net;
const classifier = knnClassifier.create();
let priceMap = {};
let logs = [];

async function setupCamera(){
  try{
    const stream = await navigator.mediaDevices.getUserMedia({video:{facingMode:'environment'}, audio:false});
    video.srcObject = stream;
    await new Promise(r => video.onloadedmetadata = r);
    status('دوربین آماده است');
  }catch(e){
    status('خطا در دسترسی به دوربین: '+e.message);
  }
}

function status(t){ statusEl.innerText = t; }

async function loadModel(){
  status('بارگذاری MobileNet...');
  net = await mobilenet.load();
  status('مدل بارگذاری شد.');
}

function captureImage(){
  const w = video.videoWidth, h = video.videoHeight;
  canvas.width=w; canvas.height=h;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(video,0,0,w,h);
  return tf.browser.fromPixels(canvas);
}

addBtn.addEventListener('click', ()=>{
  const label = labelInput.value.trim();
  const price = priceInput.value.trim();
  if(!label || !price){ alert('نام و قیمت را وارد کن'); return; }
  const img = captureImage();
  const act = net.infer(img,true);
  classifier.addExample(act,label);
  priceMap[label]=price;
  img.dispose();
  status('نمونه افزوده شد: '+label);
});

predictBtn.addEventListener('click', async ()=>{
  if(classifier.getNumClasses()===0){ alert('ابتدا چند مثال اضافه کن'); return; }
  const img = captureImage();
  const act = net.infer(img,true);
  const res = await classifier.predictClass(act,5);
  img.dispose();
  const label = res.label;
  const confidence = (res.confidences[label]*100).toFixed(1);
  predLabel.innerText='مدل: '+label;
  predPrice.innerText='قیمت: '+priceMap[label];
  predProb.innerText='اطمینان: '+confidence+'%';
  resultCard.style.display='block';
  status('تشخیص انجام شد');

  logs.push({
    مدل: label,
    قیمت: priceMap[label],
    اطمینان: confidence+"%",
    زمان: new Date().toLocaleString()
  });
});

exportBtn.addEventListener('click', ()=>{
  if(logs.length===0){ alert('هیچ رکوردی ثبت نشده'); return; }
  const ws = XLSX.utils.json_to_sheet(logs);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Results");
  XLSX.writeFile(wb, "shoes_results.xlsx");
});

(async ()=>{
  await setupCamera();
  await loadModel();
})();
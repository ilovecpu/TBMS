/const r=await apiPost({action:'clockInPhoto',row:att,photo:photo||''});/{
  a\  console.log('[DEBUG] clockIn response:',JSON.stringify(r));
}
/if(r&&!r.error){if(r.row&&r.row.photoIn)att.photoIn=r.row.photoIn;attendance.push(att);showResult('in',att.clockIn);}/{
  s/{if(r.row&&r.row.photoIn)att.photoIn=r.row.photoIn;attendance.push(att);showResult('in',att.clockIn);}/{if(r.row&&r.row.photoIn)att.photoIn=r.row.photoIn;console.log('[DEBUG] att.photoIn after clockIn:',att.photoIn);attendance.push(att);showResult('in',att.clockIn);}/
}
/else{toast('Failed to save. Try again.','error');}/{
  s/{toast(/{console.log('[DEBUG] clockIn FAILED:',JSON.stringify(r));toast(/
}

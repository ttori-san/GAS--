function addNumber(yyyy,mm){
    let numYYYY = Number(yyyy);
    let numMM = Number(mm);
    if (numMM < 12){
    numMM++;
    numYYYY;
    }else{
    numMM = 1;
    numYYYY+=1;
    }
    // let doesLess10 ;
    numMM > 10 ? numMM:numMM= `0${numMM}`;
    
    let strYYYY = String(numYYYY);
    let strMM = String(numMM);
    return [strYYYY, strMM];
}
console.log(addNumber(2021,5));
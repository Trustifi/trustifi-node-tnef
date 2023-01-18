String.prototype.replaceAll = function (search, replacement) {
    let target = this
    return target.replace(new RegExp(search, 'g'), replacement)
}

// Convert the byte array to an int value
let byteArrayToInt = ((bytes) => {
    let num = 0
    let n = 1

    for (let b in bytes) {
        num += parseInt(bytes[b]) * n
        n *= 256
    }

    return num
})

// process bytes from array using offset and length
let processBytes = ((data, offset, length) => {
    let bytesToProcess = []
    for(let i = offset; i < (offset + length); i++) {
        bytesToProcess.push(data[i])
    }

    return bytesToProcess
})

// process bytes from an array to int value using offset and length
let processBytesToInteger = ((data, offset, length) => {
    let bytesToProcess = []
    
    for(let i = offset; i < (offset + length); i++) {
        bytesToProcess.push(data[i])
    }

    return byteArrayToInt(bytesToProcess)
})

module.exports = {
    byteArrayToInt,
    processBytes,
    processBytesToInteger
}

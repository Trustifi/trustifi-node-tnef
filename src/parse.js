
const convertString = require('convert-string');
const utils = require( './util');
const mapi = require( './mapi');
const bunyan = require( 'bunyan');

const log = bunyan.createLogger({ name: 'node-tnef' })

// standard TNEF signature
const tnefSignature = 0x223e9f78
const lvlMessage = 0x01
const lvlAttachment = 0x02

// These can be used to figure out the type of attribute
// an object is
const Attribute = {
    OWNER: 0x0000, // Owner
    SENTFOR: 0x0001, // Sent For
    DELEGATE: 0x0002, // Delegate
    DATESTART: 0x0006, // Date Start
    DATEEND: 0x0007, // Date End
    AIDOWNER: 0x0008, // Owner Appointment ID
    REQUESTRES: 0x0009, // Response Requested
    FROM: 0x8000, // From
    SUBJECT: 0x8004, // Subject
    DATESENT: 0x8005, // Date Sent
    DATERECD: 0x8006, // Date Received
    MESSAGESTATUS: 0x8007, // Message Status
    MESSAGECLASS: 0x8008, // Message Class
    MESSAGEID: 0x8009, // Message ID
    PARENTID: 0x800a, // Parent ID
    CONVERSATIONID: 0x800b, // Conversation ID
    BODY: 0x800c, // Body
    PRIORITY: 0x800d, // Priority
    ATTACHDATA: 0x800f, // Attachment Data
    ATTACHTITLE: 0x8010, // Attachment File Name
    ATTACHMETAFILE: 0x8011, // Attachment Meta File
    ATTACHCREATEDATE: 0x8012, // Attachment Creation Date
    ATTACHMODIFYDATE: 0x8013, // Attachment Modification Date
    DATEMODIFY: 0x8020, // Date Modified
    ATTACHTRANSPORTFILENAME: 0x9001, // Attachment Transport File Name
    ATTACHRENDDATA: 0x9002, // Attachment Rendering Data
    MAPIPROPS: 0x9003, // MAPI Properties
    RECIPTABLE: 0x9004, // Recipients
    ATTACHMENT: 0x9005, // Attachment
    TNEFVERSION: 0x9006, // TNEF Version
    OEMCODEPAGE: 0x9007, // OEM Codepage
    ORIGNINALMESSAGECLASS: 0x9008 //Original Message Class
}

function parseBuffer(data) {
    let arr = [...data]
    return Decode(arr)
}

// right now, adds just the attachment title and data
let addAttachmentAttr = ((obj, attachment) => {
    switch (obj.Name) {
        case Attribute.ATTACHTITLE:
            attachment.Title = convertString.bytesToString(obj.Data).replaceAll('\x00', '').trim()
            break;
        case Attribute.ATTACHDATA:
            attachment.Data = obj.Data
            break;
        case Attribute.ATTACHMENT:
            let attributes = mapi.decodeMapi(obj.Data);
            if(attributes) {
                attributes.forEach(att => {
                    switch(att.Name) {
                        case mapi.MAPITypes.MAPIAttachContentId:
                            attachment.Cid = convertString.bytesToString(att.Data).replaceAll('\x00', '')
                            break;
                        case mapi.MAPITypes.MAPIAttachMimeTag:
                            attachment.ContentType = convertString.bytesToString(att.Data).replaceAll('\x00', '')
                            break;
                        case mapi.MAPITypes.MAPIAttachLongFilename:
                            attachment.LongFilename = convertString.bytesToString(att.Data).replaceAll('\x00', '')
                            break;
                        case mapi.MAPITypes.MAPIDisplayName:
                            attachment.LongFilename = convertString.bytesToString(att.Data).replaceAll('\x00', '')
                            break;
                        case mapi.MAPITypes.MAPIAttachExtension:
                            attachment.Ext = convertString.bytesToString(att.Data).replaceAll('\x00', '')
                            break;
                        case mapi.MAPITypes.MAPIAttachDataObj:
                            attachment.Content = att.Data;
                            break;
                        default:
                            if (att.Data.length && att.Data.some(itm => itm)) {
                                let attName = Object.entries(mapi.MAPITypes).find(itm => itm[1] === att.Name);
                                attName = attName && attName[0] || 'mapi_0x' + att.Name.toString(16);
                                let attValue = att.TypeSize > 0 && att.TypeSize < 8 ?
                                    utils.byteArrayToInt(att.Data) :
                                    convertString.bytesToString(att.Data).replaceAll('\x00', '');
                                attachment[attName] = {data: att.Data, value: attValue};
                            }
                    }
                })
            }
    }
})

// Decode will accept a stream of bytes in the TNEF format and extract the
// attachments and body into a Data object.
let Decode = ((data) => {

    // get the first 32 bits of the file
    let signature = utils.processBytesToInteger(data, 0, 4)

    // if the signature we get doesn't match the TNEF signature, exit
    if (signature !== tnefSignature) {
        log.warn('Value of ' + signature + ' did not equal the expected value of ' + tnefSignature)
        return null
    }

    log.info('Found a valid TNEF signature')

    // set the starting offset past the signature
    let offset = 6
    let attachment = null
    let tnef = {}
    tnef.Attachments = []

    // iterate through the data
    while (offset < data.length) {
        // get only the data within the range of offset and the array length
        let tempData = utils.processBytes(data, offset, data.length)
        // decode the TNEF objects
        let obj = decodeTNEFObject(tempData)

        if (!obj) {
            log.error('Did not get a TNEF object back, exit')
            break;
        }

        // increment offset based on the returned object's length
        offset += obj.Length

        // append attributes and attachments
        if (obj.Name === Attribute.ATTACHRENDDATA) {
            // create an empty attachment object to prepare for population
            attachment = {}
            tnef.Attachments.push(attachment)
        } else if (obj.Level === lvlAttachment) {
            // add the attachments
            addAttachmentAttr(obj, attachment)
        } else if (obj.Name === Attribute.SUBJECT) {
            tnef.Subject = convertString.bytesToString(obj.Data).replaceAll('\x00', '');
        } else if (obj.Name === Attribute.MAPIPROPS) {
            let attributes = mapi.decodeMapi(obj.Data);
            if (attributes) {
                // get the body property if it exists
                for (let attr of attributes) {
                    switch (attr.Name) {
                        case mapi.MAPITypes.MAPIBody:
                            tnef.Body = convertString.bytesToString(attr.Data).replaceAll('\x00', '');
                            break;
                        case mapi.MAPITypes.MAPIBodyHTML:
                            tnef.BodyHTML = convertString.bytesToString(attr.Data).replaceAll('\x00', '');
                            break;
                        case mapi.MAPITypes.MAPIBodyPreview:
                            tnef.BodyPreview = convertString.bytesToString(attr.Data).replaceAll('\x00', '');
                            break;
                        case mapi.MAPITypes.MAPIRtfCompressed:
                            tnef.RtfCompressed = attr.Data;
                            break;
                        case mapi.MAPITypes.MAPIPROPS:
                            tnef.MAPIPROPS = tnef.MAPIPROPS || [];
                            tnef.MAPIPROPS.push(convertString.bytesToString(attr.Data).replaceAll('\x00', ''));
                            break;
                        default:
                            if (attr.Data.length && attr.Data.some(itm => itm)) {
                                let attName = Object.entries(mapi.MAPITypes).find(itm => itm[1] === attr.Name);
                                attName = attName && attName[0] || 'mapi_0x' + attr.Name.toString(16);
                                let attValue = attr.TypeSize > 0 && attr.TypeSize < 8 ?
                                    utils.byteArrayToInt(attr.Data) :
                                    convertString.bytesToString(attr.Data).replaceAll('\x00', '');
                                tnef[attName] = {data: attr.Data, value: attValue};
                            }
                    }
                }
            }
        } else if (obj.Name === Attribute.TNEFVERSION) {
            tnef.TNEFVERSION = obj.Data;
        } else if (obj.Name === Attribute.OEMCODEPAGE) {
            tnef.OEMCODEPAGE = obj.Data;
        } else if (obj.Name === Attribute.MESSAGECLASS) {
            tnef.MESSAGECLASS = convertString.bytesToString(obj.Data).replaceAll('\x00', '');
        }
        else if(obj.Data.length && obj.Data.some(itm => itm)) {
            let attName = Object.entries(Attribute).find(itm => itm[1] === obj.Name);
            attName = attName && attName[0] || '0x' + obj.Name.toString(16);
            tnef[attName] = {data: obj.Data, value: convertString.bytesToString(obj.Data).replaceAll('\x00', '')};
        }
    }

    // return the final TNEF object
    return tnef;
})

let decodeTNEFObject = ((data) => {
    let tnefObject = {}
    let offset = 0
    let object = {}

    // level
    object.Level = utils.processBytesToInteger(data, offset, 1)
    offset++

    // name
    object.Name = utils.processBytesToInteger(data, offset, 2)
    offset += 2

    // type
    object.Type = utils.processBytesToInteger(data, offset, 2)
    offset += 2

    // attribute length
    let attLength = utils.processBytesToInteger(data, offset, 4)
    offset += 4

    // data
    if(data.length > offset+attLength) {
        object.Data = utils.processBytes(data, offset, attLength)
        offset += attLength
    }
    else {
        object.Data = []
    }

    offset += 2

    // length
    object.Length = offset

    return object
})

module.exports = {
    parseBuffer
};

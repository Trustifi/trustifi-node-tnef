
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
    ATTOWNER: 0x0000, // Owner
    ATTSENTFOR: 0x0001, // Sent For
    ATTDELEGATE: 0x0002, // Delegate
    ATTDATESTART: 0x0006, // Date Start
    ATTDATEEND: 0x0007, // Date End
    ATTAIDOWNER: 0x0008, // Owner Appointment ID
    ATTREQUESTRES: 0x0009, // Response Requested
    ATTFROM: 0x8000, // From
    ATTSUBJECT: 0x8004, // Subject
    ATTDATESENT: 0x8005, // Date Sent
    ATTDATERECD: 0x8006, // Date Received
    ATTMESSAGESTATUS: 0x8007, // Message Status
    ATTMESSAGECLASS: 0x8008, // Message Class
    ATTMESSAGEID: 0x8009, // Message ID
    ATTPARENTID: 0x800a, // Parent ID
    ATTCONVERSATIONID: 0x800b, // Conversation ID
    ATTBODY: 0x800c, // Body
    ATTPRIORITY: 0x800d, // Priority
    ATTATTACHDATA: 0x800f, // Attachment Data
    ATTATTACHTITLE: 0x8010, // Attachment File Name
    ATTATTACHMETAFILE: 0x8011, // Attachment Meta File
    ATTATTACHCREATEDATE: 0x8012, // Attachment Creation Date
    ATTATTACHMODIFYDATE: 0x8013, // Attachment Modification Date
    ATTDATEMODIFY: 0x8020, // Date Modified
    ATTATTACHTRANSPORTFILENAME: 0x9001, // Attachment Transport File Name
    ATTATTACHRENDDATA: 0x9002, // Attachment Rendering Data
    ATTMAPIPROPS: 0x9003, // MAPI Properties
    ATTRECIPTABLE: 0x9004, // Recipients
    ATTATTACHMENT: 0x9005, // Attachment
    ATTTNEFVERSION: 0x9006, // TNEF Version
    ATTOEMCODEPAGE: 0x9007, // OEM Codepage
    ATTORIGNINALMESSAGECLASS: 0x9008 //Original Message Class
}

function parseBuffer(data) {
    let arr = [...data]
    return Decode(arr)
}

// right now, adds just the attachment title and data
let addAttachmentAttr = ((obj, attachment) => {
    switch (obj.Name) {
        case Attribute.ATTATTACHTITLE:
            attachment.Title = convertString.bytesToString(obj.Data).replaceAll('\x00', '').trim();
            break;
        case Attribute.ATTATTACHDATA:
            attachment.Data = obj.Data
            break;
        case Attribute.ATTATTACHMENT:
            let attributes = mapi.decodeMapi(obj.Data);
            if(attributes) {
                attributes.forEach(att => {
                    switch(att.Name) {
                        case mapi.MAPITypes.MAPIAttachContentId:
                            attachment.Cid = convertString.bytesToString(att.Data).replaceAll('\x00', '')
                            break;
                        case mapi.MAPITypes.MAPIAttachMimeTag:
                            attachment.ContentType = attachment.ContentType || convertString.bytesToString(att.Data).replaceAll('\x00', '');
                            break;
                        case mapi.MAPITypes.MAPIAttachTag:
                            attachment.Tag = att.Data[8] === 1 ? 'TNEF' : att.Data[8] === 3 ? 'OLE' : att.Data[8] === 4 ? 'Mime' : 'Unknown';
                            if (att.Data[8] === 3) {
                                if (att.Data[9]) attachment.Tag += att.Data[9];
                                if (att.Data[10] === 1) attachment.Tag += ' storage';
                            }
                            break;
                        case mapi.MAPITypes.MAPIAttachLongFilename:
                            attachment.LongFilename = convertString.bytesToString(att.Data).replaceAll('\x00', '')
                            break;
                        case mapi.MAPITypes.MAPIAttachExtension:
                            attachment.Ext = convertString.bytesToString(att.Data).replaceAll('\x00', '')
                            break;
                        case mapi.MAPITypes.MAPIAttachDataObj:
                            let signature = utils.processBytesToInteger(att.Data.slice(16), 0, 4);
                            if (signature === tnefSignature) {
                                attachment.Content = att.Data.slice(16);
                                attachment.ContentType = 'application/ms-tnef';
                            } else {
                                attachment.Content = att.Data;
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
        if (obj.Name === Attribute.ATTATTACHRENDDATA) {
            // create an empty attachment object to prepare for population
            attachment = {}
            tnef.Attachments.push(attachment)
        } else if (obj.Level === lvlAttachment) {
            // add the attachments
            addAttachmentAttr(obj, attachment)
        } else if (obj.Name === Attribute.ATTSUBJECT) {
            tnef.Subject = obj.Data;
        } else if (obj.Name === Attribute.ATTMAPIPROPS) {
            let attributes = mapi.decodeMapi(obj.Data);
            if (attributes) {
                // get the body property if it exists
                for (let attr of attributes) {
                    switch (attr.Name) {
                        case mapi.MAPITypes.MAPIBody:
                            tnef.Body = attr.Data;
                            break;
                        case mapi.MAPITypes.MAPIBodyHTML:
                            tnef.BodyHTML = attr.Data;
                            break;
                        case mapi.MAPITypes.MAPIBodyPreview:
                            tnef.BodyPreview = attr.Data;
                            break;
                        case mapi.MAPITypes.MAPIRtfCompressed:
                            tnef.RtfCompressed = attr.Data;
                            break;
                        case mapi.MAPITypes.MAPIMessageClass:
                            tnef.MessageClass = convertString.bytesToString(attr.Data).replaceAll('\x00', '');
                            break;
                        default:
                            /*if (attr.Data) {      //DEBUG
                                let name = Object.entries(mapi.MAPITypes).find(a => a[1] === attr.Name);
                                let data = convertString.bytesToString(attr.Data).replaceAll('\x00', '').trim();
                                if (data) {
                                    tnef.mapi = tnef.mapi || {};
                                    tnef.mapi[name && name[0] || attr.Name] = data;
                                }
                            }*/
                    }
                }
            }
        }
        /*else if (obj.Data) {      //DEBUG
            let name = Object.entries(Attribute).find(a => a[1] === obj.Name);
            let data = convertString.bytesToString(obj.Data).replaceAll('\x00', '').trim();
            if (data) {
                tnef.extra = tnef.extra || {};
                tnef.extra[name && name[0] || obj.Name] = data;
            }
        }*/
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

const ftp = require("basic-ftp")

async function getAllProducts() {
    const client = new ftp.Client()
    client.ftp.verbose = true
    try {
        await client.access({
            host: "192.168.88.43",
            user: "marjan-ftp",
            password: "Enter_paswd_77",
            secure: false
        })
        await client.downloadTo("./source/All.xlsx", "All.xlsx")
    } catch (err) {
        console.log(err)
    }
    client.close()
}

module.exports = getAllProducts
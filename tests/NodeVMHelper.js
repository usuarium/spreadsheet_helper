const path = require('path')
const fs = require('fs')
const util = require('util')
const vm = require('vm')

fs.readdirAsync = util.promisify(fs.readdir)
fs.readFileAsync = util.promisify(fs.readFile)
fs.statAsync = util.promisify(fs.stat)

class NodeVMHelper {
    constructor() {
        this.ctx = null
    }
    
    async readJavascriptFiles(dir, ext) {
        let results = []
        let directoryContent = await fs.readdirAsync(dir)
        
        if (ext === undefined) {
            ext = '.js'
        }
        
        for (let filename of directoryContent) {
            let fullPath = path.join(dir, filename)
            let stat = await fs.statAsync(fullPath)
            
            if (stat && stat.isDirectory()) {
                results = results.concat(await this.readJavascriptFiles(fullPath, ext))
            }
            else if (path.extname(fullPath) == ext) {
                results.push(fullPath)
            }
        }
        
        return results
    }
    
    async loadContext(directory, globals) {
        this.ctx = vm.createContext(globals)

        await this.addContext(this.ctx, directory, globals)
    }
    
    async addContext(ctx, directory, globals) {
        globals = globals || {}
        let files = await this.readJavascriptFiles(directory)
        
        for (let file of files) {
            let content = await fs.readFileAsync(file)
            
            try {
                vm.runInContext(content, ctx, file)
            }
            catch (error) {
                console.error(error)
            }
        }
    }
    
    instantiateClass(classname) {
        if (this.ctx[`${classname}_factory`] === undefined) {
            let content = `function ${classname}_factory() {
                return new ${classname}()
            }`
            vm.runInContext(content, this.ctx, 'main.js')
        }
        return this.ctx[`${classname}_factory`]()
    }
}

module.exports = NodeVMHelper
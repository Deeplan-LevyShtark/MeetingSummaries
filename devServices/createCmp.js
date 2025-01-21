const fs = require('fs');
const path = require('path');
const projectConfig = require('../config/config.json');

// Docs:
// quick function cmp: npm run createCmp -- -n <name>
// quick class cmp: npm run createCmp -- -n <name> -ct class
// dir cmp: npm run createCmp -- -n <name> -d components/Test/Test1/Test2...
// full cmp example: npm run createCmp -- -n <name> [-d <dir>] [-ct <cmpType>]

class Log {
    static colors = {
        reset: "\x1b[0m",
        red: "\x1b[31m",
        green: "\x1b[32m",
        yellow: "\x1b[33m",
    };

    // Method for logging success messages
    static success(message) {
        console.log(`${this.colors.green}%s${this.colors.reset}`, message);
    }

    // Method for logging error messages
    static error(message) {
        console.log(`${this.colors.red}%s${this.colors.reset}`, message);
    }

    // Method for logging warning messages
    static warning(message) {
        console.log(`${this.colors.yellow}%s${this.colors.reset}`, message);
    }
}

const validOptions = {
    name: 'name',
    n: 'name',
    dir: 'dir',
    d: 'dir',
    cmpType: 'cmpType',
    ct: 'cmpType',
    style: 'style',
    s: 'style',
}

const defaultOptions = {
    dir: 'components',
    style: 'scss',
    cmpType: 'func'
}

const extractProjectName = () => {
    const projectNameLocation = Object.keys(projectConfig.localizedResources)[0];
    const projectNameStr = projectConfig.localizedResources[projectNameLocation];

    return projectNameStr.split('/')[2];
}

const componentClassTemplate = (cmpName, style) => {
    const name = cmpName.charAt(0).toLowerCase() + cmpName.slice(1);

    return (`import * as React from 'react';
import ${style === 'scss' ? 'styles from ' : ''}'./${cmpName}${style === 'scss' ? '.module' : ''}.${style}';

export interface ${cmpName}Props {

}

export interface ${cmpName}States {

}

export default class ${cmpName} extends React.Component<${cmpName}Props, ${cmpName}States> {

  constructor(props: ${cmpName}Props) {
    super(props);
    this.state = {

    }
  }

  public render(): React.ReactElement<${cmpName}Props> {

    return (
      <div className=${style === 'scss' ? `{styles.${name}Container}` : `"${name}Container"`}>

      </div>
    );
  }
}`)
}

const componentFunctionTemplate = (cmpName, style) => {
    const name = cmpName.charAt(0).toLowerCase() + cmpName.slice(1);

    return (`import * as React from 'react';
import ${style === 'scss' ? 'styles from ' : ''}'./${cmpName}${style === 'scss' ? '.module' : ''}.${style}';
    
interface ${cmpName}Props {
        
}
    
export default function ${cmpName}({ }: ${cmpName}Props) {
    
    return (
        <div className=${style === 'scss' ? `{styles.${name}Container}` : `"${name}Container"`}>
        
        </div>
    );
}`)
}

const styleTemplate = (styleName) => {
    const name = styleName.charAt(0).toLowerCase() + styleName.slice(1);
    return `.${name}Container {
    width: 100%;
    height: 100%;
}
`
}

const showHelp = () => {
    console.log('createCmp usage:')
    console.log('node createCmp -name <name> [-dir <dir>] [-cmpType <cmpType>] [-style <css | scss>]')
    console.log('')
    console.log('   -name      required')
    console.log('   -dir       optional, default: components')
    console.log('   -cmpType   optional, default: func')
    console.log('   -style     optional, default: scss')
}

// removes '-' and '--' from option name
const cleanOption = (op) => {
    if (op.includes('--')) {
        op = op.substr(2, op.length);
        return validOptions[op];
    } else if (op.includes('-')) {
        op = op.substr(1, op.length);
        return validOptions[op];
    };

    throw "option name must start with '--' or '-";
}

// parse and validate options and values
const parseOptions = (args) => {
    const options = Object.assign({}, defaultOptions)
    // check for non-char
    const regExp = /^[a-zA-Z]+$/;

    for (let i = 0; i < args.length; i = i + 2) {
        // remove '--' or '-' from option name
        const op = cleanOption(args[i])
        // get option value
        const val = args[i + 1]

        // validate option
        if (!validOptions[op]) throw `'${op}' is not a valid option!`;
        // validate value
        if (val === null || val === undefined || val === '') throw 'each option must have value!';
        // validate cmpType
        if ((op === 'cmpType' || op === 'ct') && !(val === 'class' || val === 'func')) throw 'option cmpType can be `class` or `func` only!';
        // validate style
        if ((op === 'style' || op === 's') && !(val === 'css' || val === 'scss')) throw 'option style can be `css` or `scss` only!';
        // validate value
        if (!regExp.test(val) && (op === 'cmpType' || op === 'ct')) throw 'options value must be chars only!';

        options[op] = val;
    }

    if (!options.name) throw 'name is required!';

    return options;
}

try {
    // get args
    const args = process.argv.splice(2);

    // check for help
    if (args[0] === '--help' || args[0] === '-help' || args[0] === '--h' || args[0] === '-h' || args[0] === '--') {
        showHelp();
        return;
    }

    // parse options
    const options = parseOptions(args);
    const { name, dir, style, cmpType } = options;
    // build path
    const projectTypePath = path.resolve(__dirname, '..', 'src', 'webparts');
    const projectType = fs.existsSync(projectTypePath) ? 'webparts' : 'extensions';
    const projectName = extractProjectName();
    const rootDir = path.join(__dirname, '..', `/src/${projectType}/${projectName}/${dir}/${name}`);

    // create dir
    fs.mkdirSync(rootDir);
    // write styles
    fs.writeFileSync(`${rootDir}/${name}${style === 'scss' ? '.module' : ''}.${style}`, styleTemplate(name));

    // select class or function template
    const cmpTemplate = cmpType === 'class' ? componentClassTemplate : componentFunctionTemplate;
    // write template
    fs.writeFileSync(`${rootDir}/${name}.cmp.tsx`, cmpTemplate(name, style));
    // notify user
    Log.success(`Component '${name}' created`);
    Log.success(`Full path: ${rootDir}`);
} catch (e) {
    Log.error('ERROR: ' + (e.message ? e.message : e));
}
import { SPComponentLoader } from "@microsoft/sp-loader";
import JQueryStatic = require('jquery');
declare var window: any;
declare var jQuery: JQueryStatic;
declare var $: JQueryStatic;

namespace CatDsgSpWp1006jQueryLoader {
    export function LoadDependencies(JqueryScript: string, dependencies: string[]): Promise<object> {
        if ('undefined' != typeof window.jQuery) {
            jQuery = window.jQuery as JQueryStatic as JQueryStatic;
            return Load(dependencies);
        } else {
            return SPComponentLoader.loadScript(JqueryScript, { globalExportsName: 'jQuery' }).then(() => {
                return Load(dependencies);
            });
        }
    }
    function Load(dependencies: string[]): Promise<object> {
        var scripts: Promise<object>[] = [];
        dependencies.forEach(depenency => {
            scripts.push(SPComponentLoader.loadScript(depenency));
        });

        return Promise.all(scripts);
    }
}
export default CatDsgSpWp1006jQueryLoader;
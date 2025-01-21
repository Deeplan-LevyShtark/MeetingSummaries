import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/content-types";
import "@pnp/sp/site-users/web";
// import "@pnp/sp/items/get-all";
import "@pnp/sp/profiles"

// import { graphfi, SPFx as graphSPFx } from "@pnp/graph";
// import { GraphFI } from "@pnp/graph/fi";
// import "@pnp/graph/teams";
// import "@pnp/graph/teams";
// import "@pnp/graph/planner";
// import "@pnp/graph/users";
// import "@pnp/graph/messages";

// SP:
// npm install @pnp/sp @pnp/nodejs --save

// Graph:
// npm install @pnp/core @pnp/queryable @pnp/graph --save

var sp: SPFI;
// var _graph: GraphFI;

export const getSP = (context?: WebPartContext): SPFI => {
    if (!sp && context) sp = spfi().using(SPFx(context));

    return sp;
};

// export const getGraph = (context?: WebPartContext): GraphFI => {
//     if (!_graph && context) _graph = graphfi().using(graphSPFx(context));

//     return _graph;
// };
import { SPFI } from "@pnp/sp";
import { IItems } from "./IItems";

export interface ITreinamentoState {
    items: Array<IItems>;
    contador: number;
    sp: SPFI; // instância do PnP SP para interações com o SharePoint
}
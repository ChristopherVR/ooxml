import * as xpath from 'xpath';
import { xpathSelector } from './xpath-selector';

export type { XmlXPathOptions } from './xpath-selector';

/** XPath 1.0 node selection; invalid expressions and scalar results throw. */
export const selectXmlNodes = xpathSelector(xpath);

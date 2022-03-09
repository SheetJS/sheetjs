import { format, load, get_table, load_table, parse_date_code, is_date, SSF$Table, SSF$Date } from 'ssf';

const t1: string = format("General", 123.456);
const t2: string = format(0, 234.567);
const t3: string = format("@", "1234.567");

load('"This is "0.00', 70);
load('"This is "0');

const tbl: SSF$Table = get_table();
load_table(tbl);

const date: SSF$Date = parse_date_code(43150);
const sum: number = date.D + date.T + date.y + date.m + date.d + date.H + date.M + date.S + date.q + date.u;

const isdate: boolean = is_date("YYYY-MM-DD");

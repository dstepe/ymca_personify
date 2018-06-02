create table products (
  product_code text not null,
  branch text not null,
  type text not null,
  description text not null,
  summary text,
  session text
);

create index products_product_code on products (product_code);
create index products_branch on products (branch);
create index products_type on products (type);
create index products_description on products (description);
create index products_summary on products (summary);
create index products_session on products (session);

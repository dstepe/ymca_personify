create table companies (
  t_id text not null,
  p_id text not null,
  c_name text
);

create index companies_t_id on companies (t_id);
create index companies_p_id on companies (p_id);
create index companies_name on companies (c_name);

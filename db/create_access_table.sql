create table access (
  t_id text not null,
  p_id text not null,
  access text,
  reason text
);

create index ids_t_id on ids (t_id);
create index ids_p_id on ids (p_id);

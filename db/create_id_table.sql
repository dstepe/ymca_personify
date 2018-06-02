create table ids (
  t_id text not null,
  p_id text not null
);

create index ids_t_id on ids (t_id);

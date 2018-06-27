create table ids (
  t_id text not null,
  p_id text not null
);

create unique index ids_t_id on ids (t_id);
create unique index ids_p_id on ids (p_id);

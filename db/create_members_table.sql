create table members (
  p_id text not null,
  f_id text not null,
  membership text,
  b_id text not null,
  is_primary text not null
);

create unique index members_p_id on members (p_id);
create index members_f_id on members (f_id);

create table name_map (
  p_id text not null,
  c_name text not null
);

create index name_map_p_id on name_map (p_id);
create index name_map_name_id on name_map (c_name);

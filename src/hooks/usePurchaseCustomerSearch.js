import { useEffect, useMemo, useRef, useState } from "react";
import { sb } from "../lib/supabaseClient";

const MIN_SEARCH_LENGTH = 2;

export function usePurchaseCustomerSearch({
  search,
  orders,
  debounceMs = 250,
  queryBuilder,
  postFilter
}) {
  const [searchResults, setSearchResults] = useState([]);
  const [searchLoading, setSearchLoading] = useState(false);
  const requestSeqRef = useRef(0);
  const lastFetchKeyRef = useRef("");
  const queryBuilderRef = useRef(queryBuilder);
  const postFilterRef = useRef(postFilter);

  useEffect(() => {
    queryBuilderRef.current = queryBuilder;
  }, [queryBuilder]);

  useEffect(() => {
    postFilterRef.current = postFilter;
  }, [postFilter]);

  const orderNameById = useMemo(() => {
    const map = new Map();
    (orders || []).forEach((order) => {
      map.set(String(order.id), order.orderName || "طلبية");
    });
    return map;
  }, [orders]);

  useEffect(() => {
    const query = String(search || "").trim();
    if (query.length < MIN_SEARCH_LENGTH) {
      requestSeqRef.current += 1;
      lastFetchKeyRef.current = "";
      setSearchResults([]);
      setSearchLoading(false);
      return;
    }

    const orderIds = (orders || []).map((order) => order.id);
    if (!orderIds.length) {
      requestSeqRef.current += 1;
      lastFetchKeyRef.current = "";
      setSearchResults([]);
      setSearchLoading(false);
      return;
    }

    const orderIdsKey = orderIds.map((id) => String(id)).sort().join(",");
    const fetchKey = `${query}__${orderIdsKey}`;
    if (lastFetchKeyRef.current === fetchKey) {
      return;
    }
    lastFetchKeyRef.current = fetchKey;

    const requestSeq = ++requestSeqRef.current;
    const timer = window.setTimeout(async () => {
      setSearchLoading(true);
      try {
        let request = sb
          .from("purchases")
          .select("id, order_id, customer_name, price, created_at, pickup_point")
          .in("order_id", orderIds)
          .eq("collected", false)
          .ilike("customer_name", `%${query}%`)
          .order("created_at", { ascending: false })
          .limit(50);

        if (typeof queryBuilderRef.current === "function") {
          request = queryBuilderRef.current(request, { query, orderIds }) || request;
        }

        const { data, error } = await request;
        if (error) throw error;
        if (requestSeqRef.current !== requestSeq) return;

        let list = data || [];
        if (typeof postFilterRef.current === "function") {
          list = list.filter((item) => postFilterRef.current(item, { query, orderIds }));
        }
        if (requestSeqRef.current !== requestSeq) return;

        setSearchResults(
          list.map((item) => ({
            ...item,
            orderName: orderNameById.get(String(item.order_id)) || "طلبية"
          }))
        );
      } catch (error) {
        console.error(error);
        if (requestSeqRef.current === requestSeq) {
          setSearchResults([]);
        }
      } finally {
        if (requestSeqRef.current === requestSeq) {
          setSearchLoading(false);
        }
      }
    }, debounceMs);

    return () => window.clearTimeout(timer);
  }, [debounceMs, orderNameById, orders, search]);

  const clearSearchResults = () => setSearchResults([]);

  return { searchResults, searchLoading, clearSearchResults };
}

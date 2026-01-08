import { supabase } from "../supabase_client.js";

const SIGNED_URL_TTL = 60 * 60 * 24;

function resolvePhotoExtension(file) {
  const nameExt = file?.name?.split(".").pop();
  if (nameExt && nameExt !== file.name) {
    return nameExt.toLowerCase();
  }
  const typeExt = file?.type?.split("/").pop();
  if (typeExt) {
    return typeExt.toLowerCase();
  }
  return "jpg";
}

function warnSignedUrlFailure(error) {
  console.error("Failed to create signed URL.", error?.message || error);
  if (typeof globalThis?.showSnackbar === "function") {
    globalThis.showSnackbar("Some photos could not be loaded. Please retry.");
  }
}

export async function listCustomerPhotos(customerId) {
  if (!customerId) {
    throw new Error("Missing customer id");
  }
  const { data, error } = await supabase
    .from("customer_photos")
    .select("*")
    .eq("customer_id", customerId)
    .order("created_at", { ascending: false });
  if (error) {
    throw error;
  }

  const rows = data || [];
  const signedRows = await Promise.all(
    rows.map(async (row) => {
      const { data: signed, error: signedError } = await supabase.storage
        .from("customer-photos")
        .createSignedUrl(row.image_path, SIGNED_URL_TTL);
      if (signedError) {
        warnSignedUrlFailure(signedError);
        return { ...row, signedUrl: null };
      }
      return { ...row, signedUrl: signed?.signedUrl || null };
    })
  );

  return signedRows;
}

export async function uploadCustomerPhoto({ customerId, file, caption, store_location }) {
  if (!customerId) {
    throw new Error("Missing customer id");
  }
  if (!file) {
    throw new Error("Please choose a photo file");
  }

  const ext = resolvePhotoExtension(file);
  const path = `${customerId}/${crypto.randomUUID()}.${ext}`;

  const { error: uploadError } = await supabase.storage
    .from("customer-photos")
    .upload(path, file, { contentType: file.type, upsert: false });
  if (uploadError) {
    throw uploadError;
  }

  const payload = {
    customer_id: customerId,
    image_path: path,
    caption: caption?.trim() || null,
    store_location: store_location?.trim() || null,
  };
  const { data, error: insertError } = await supabase
    .from("customer_photos")
    .insert(payload)
    .select("*")
    .single();
  if (insertError) {
    try {
      await supabase.storage.from("customer-photos").remove([path]);
    } catch (cleanupError) {
      console.error("Failed to cleanup customer photo upload.", cleanupError);
    }
    throw insertError;
  }

  return data;
}

export async function updateCustomerPhoto({ id, caption, store_location }) {
  if (!id) {
    throw new Error("Missing photo id");
  }
  const { data, error } = await supabase
    .from("customer_photos")
    .update({
      caption: caption?.trim() || null,
      store_location: store_location?.trim() || null,
    })
    .eq("id", id)
    .select("*")
    .single();
  if (error) {
    throw error;
  }
  return data;
}

export async function deleteCustomerPhoto({ id }) {
  if (!id) {
    throw new Error("Missing photo id");
  }
  const { data, error } = await supabase
    .from("customer_photos")
    .select("id, image_path")
    .eq("id", id)
    .single();
  if (error) {
    throw error;
  }
  if (data?.image_path) {
    await supabase.storage.from("customer-photos").remove([data.image_path]);
  }
  const { error: deleteError } = await supabase.from("customer_photos").delete().eq("id", id);
  if (deleteError) {
    throw deleteError;
  }
}
